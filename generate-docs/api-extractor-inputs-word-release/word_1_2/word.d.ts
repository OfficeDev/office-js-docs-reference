import { OfficeExtension } from "../../api-extractor-inputs-office/office"
import { Office as Outlook} from "../../api-extractor-inputs-outlook/outlook"
////////////////////////////////////////////////////////////////
/////////////////////// Begin Word APIs ////////////////////////
////////////////////////////////////////////////////////////////

export declare namespace Word {
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
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
         * Gets the collection of rich text content control objects in the body.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly contentControls: Word.ContentControlCollection;
        
        
        /**
         * Gets the text format of the body. Use this to get and set font name, size, color, and other properties.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly font: Word.Font;
        
        /**
         * Gets the collection of `InlinePicture` objects in the body. The collection doesn't include floating images.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly inlinePictures: Word.InlinePictureCollection;
        
        /**
         * Gets the collection of `Paragraph` objects in the body.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * Important: Paragraphs in tables aren't returned for requirement sets 1.1 and 1.2. From requirement set 1.3, paragraphs in tables are also returned.
         */
        readonly paragraphs: Word.ParagraphCollection;
        
        
        /**
         * Gets the content control that contains the body. Throws an `ItemNotFound` error if there isn't a parent content control.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly parentContentControl: Word.ContentControl;
        
        
        
        
        
        /**
         * Specifies the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        style: string;
        
        /**
         * Gets the text of the body. Use the `insertText` method to insert text.
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
        set(properties: Interfaces.BodyUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Word.Body): void;
        /**
         * Clears the contents of the `Body` object. The user can perform the undo operation on the cleared content.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        clear(): void;
        
        
        /**
         * Gets an HTML representation of the `Body` object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method doesn't return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use `Body.getOoxml()` and convert the returned XML to HTML.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        /**
         * Gets the OOXML (Office Open XML) representation of the `Body` object.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        
        
        
        
        /**
         * Inserts a break at the specified location in the main document.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param breakType - The break type to add to the body.
         * @param insertLocation - The value must be `start` or `end`.
         */
        insertBreak(breakType: Word.BreakType | "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | "Start" | "End"): void;
        /**
         * Wraps the `Body` object with a content control.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * Note: The `contentControlType` parameter was introduced in WordApi 1.5. `plainText` support was added in WordApi 1.5. `checkBox` support was added in WordApi 1.7.
         * `dropDownList` and `comboBox` support was added in WordApi 1.9. Support for `buildingBlockGallery`, `datePicker`, `group`, `picture`, and `repeatingSection` was added in WordApiDesktop 1.3.
         *
         * @param contentControlType - Optional. Content control type to insert. Must be `richText`, `plainText`, `checkBox`, `dropDownList`, `comboBox`, `buildingBlockGallery`, `datePicker`, `repeatingSection`, `picture`, or `group`. The default is `richText`.
         */
        insertContentControl(contentControlType?: Word.ContentControlType.richText | Word.ContentControlType.plainText | Word.ContentControlType.checkBox | Word.ContentControlType.dropDownList | Word.ContentControlType.comboBox | Word.ContentControlType.buildingBlockGallery | Word.ContentControlType.datePicker | Word.ContentControlType.repeatingSection | Word.ContentControlType.picture | Word.ContentControlType.group | "RichText" | "PlainText" | "CheckBox" | "DropDownList" | "ComboBox" | "BuildingBlockGallery" | "DatePicker" | "RepeatingSection" | "Picture" | "Group"): Word.ContentControl;
        /**
         * Inserts a document into the body at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * Insertion isn't supported if the document being inserted contains an ActiveX control (likely in a form field). Consider replacing such a form field with a content control or other option appropriate for your scenario.
         *
         * @param base64File - The Base64-encoded content of a .docx file.
         * @param insertLocation - The value must be `replace`, `start`, or `end`.
         */
        insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Inserts HTML at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param html - The HTML to be inserted in the document.
         * @param insertLocation - The value must be `replace`, `start`, or `end`.
         */
        insertHtml(html: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Inserts a picture into the body at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - The Base64-encoded image to be inserted in the body.
         * @param insertLocation - The value must be `start` or `end`.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | "Start" | "End"): Word.InlinePicture;
        /**
         * Inserts OOXML at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - The OOXML to be inserted.
         * @param insertLocation - The value must be `replace`, `start`, or `end`.
         */
        insertOoxml(ooxml: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Inserts a paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - The paragraph text to be inserted.
         * @param insertLocation - The value must be `start` or `end`.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | "Start" | "End"): Word.Paragraph;
        
        /**
         * Inserts text into the body at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param text - Text to be inserted.
         * @param insertLocation - The value must be `replace`, `start`, or `end`.
         */
        insertText(text: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Performs a search with the specified search options on the scope of the `Body` object. The search results are a collection of `Range` objects.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param searchText - The search text. Can be a maximum of 255 characters.
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
         * @param selectionMode - Optional. The selection mode must be `select`, `start`, or `end`. `select` is the default.
         */
        select(selectionMode?: Word.SelectionMode): void;
        /**
         * Selects the body and navigates the Word UI to it.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param selectionMode - Optional. The selection mode must be `select`, `start`, or `end`. `select` is the default.
         */
        select(selectionMode?: "Select" | "Start" | "End"): void;
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
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.Body;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.Body;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Word.Body` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.BodyData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.BodyData;
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    /**
     * Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text, plain text, checkbox, dropdown list, and combo box content controls are supported.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    export class ContentControl extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        
        
        /**
         * Gets the collection of `ContentControl` objects in the content control.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly contentControls: Word.ContentControlCollection;
        
        
        
        
        /**
         * Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly font: Word.Font;
        
        
        /**
         * Gets the collection of `InlinePicture` objects in the content control. The collection doesn't include floating images.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly inlinePictures: Word.InlinePictureCollection;
        
        /**
         * Gets the collection of `Paragraph` objects in the content control.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * Important: For requirement sets 1.1 and 1.2, paragraphs in tables wholly contained within this content control aren't returned. From requirement set 1.3, paragraphs in such tables are also returned.
         */
        readonly paragraphs: Word.ParagraphCollection;
        
        /**
         * Gets the content control that contains the content control. Throws an `ItemNotFound` error if there isn't a parent content control.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly parentContentControl: Word.ContentControl;
        
        
        
        
        
        
        
        
        
        /**
         * Specifies the appearance of the content control. The value can be `boundingBox`, `tags`, or `hidden`.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        appearance: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden";
        /**
         * Specifies a value that indicates whether the user can delete the content control. Mutually exclusive with `removeWhenEdited`.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        cannotDelete: boolean;
        /**
         * Specifies a value that indicates whether the user can edit the contents of the content control.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        cannotEdit: boolean;
        /**
         * Specifies the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        color: string;
        /**
         * Gets an integer that represents the content control identifier.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly id: number;
        /**
         * Specifies the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        placeholderText: string;
        /**
         * Specifies a value that indicates whether the content control is removed after it's edited. Mutually exclusive with `cannotDelete`.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        removeWhenEdited: boolean;
        /**
         * Specifies the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the `styleBuiltIn` property.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        style: string;
        
        
        /**
         * Specifies a tag to identify a content control.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        tag: string;
        /**
         * Gets the text of the content control.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly text: string;
        /**
         * Specifies the title for a content control.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        title: string;
        /**
         * Gets the content control type. Only rich text, plain text, check box, dropdown list, combo box, building block gallery, date picker, repeating section, picture, and group content controls are supported currently.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly type: Word.ContentControlType | "Unknown" | "RichTextInline" | "RichTextParagraphs" | "RichTextTableCell" | "RichTextTableRow" | "RichTextTable" | "PlainTextInline" | "PlainTextParagraph" | "Picture" | "BuildingBlockGallery" | "CheckBox" | "ComboBox" | "DropDownList" | "DatePicker" | "RepeatingSection" | "RichText" | "PlainText" | "Group";
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
         * Deletes the content control and its content. If `keepContent` is set to `true`, the content isn't deleted.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param keepContent - Indicates whether the content should be deleted with the content control. If `keepContent` is set to `true`, the content isn't deleted.
         */
        delete(keepContent: boolean): void;
        
        
        /**
         * Gets an HTML representation of the `ContentControl` object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method doesn't return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use `ContentControl.getOoxml()` and convert the returned XML to HTML.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        /**
         * Gets the Office Open XML (OOXML) representation of the `ContentControl` object.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        
        
        
        
        
        /**
         * Inserts a break at the specified location in the main document. This method cannot be used with `richTextTable`, `richTextTableRow`, and `richTextTableCell` content controls.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param breakType - Type of break.
         * @param insertLocation - The value must be `start`, `end`, `before`, or `after`.
         */
        insertBreak(breakType: Word.BreakType | "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | Word.InsertLocation.before | Word.InsertLocation.after | "Start" | "End" | "Before" | "After"): void;
        /**
         * Inserts a document into the content control at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * Insertion isn't supported if the document being inserted contains an ActiveX control (likely in a form field). Consider replacing such a form field with a content control or other option appropriate for your scenario.
         *
         * @param base64File - The Base64-encoded content of a .docx file.
         * @param insertLocation - The value must be `replace`, `start`, or `end`. `replace` cannot be used with `richTextTable` and `richTextTableRow` content controls.
         */
        insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Inserts HTML into the content control at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param html - The HTML to be inserted in to the content control.
         * @param insertLocation - The value must be `replace`, `start`, or `end`. `replace` cannot be used with `richTextTable` and `richTextTableRow` content controls.
         */
        insertHtml(html: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Inserts an inline picture into the content control at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - The Base64-encoded image to be inserted in the content control.
         * @param insertLocation - The value must be `replace`, `start`, or `end`. `replace` cannot be used with `richTextTable` and `richTextTableRow` content controls.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.InlinePicture;
        /**
         * Inserts OOXML into the content control at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - The OOXML to be inserted in to the content control.
         * @param insertLocation - The value must be `replace`, `start`, or `end`. `replace` cannot be used with `richTextTable` and `richTextTableRow` content controls.
         */
        insertOoxml(ooxml: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Inserts a paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - The paragraph text to be inserted.
         * @param insertLocation - The value must be `start`, `end`, `before`, or `after`. `before` and `after` cannot be used with `richTextTable`, `richTextTableRow`, and `richTextTableCell` content controls.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | Word.InsertLocation.before | Word.InsertLocation.after | "Start" | "End" | "Before" | "After"): Word.Paragraph;
        
        /**
         * Inserts text into the content control at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param text - The text to be inserted in to the content control.
         * @param insertLocation - The value must be `replace`, `start`, or `end`. `replace` cannot be used with `richTextTable` and `richTextTableRow` content controls.
         */
        insertText(text: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Performs a search with the specified search options on the scope of the `ContentControl` object. The search results are a collection of `Range` objects.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param searchText - The search text.
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
         * @param selectionMode - Optional. The selection mode must be `select`, `start`, or `end`. `select` is the default.
         */
        select(selectionMode?: Word.SelectionMode): void;
        /**
         * Selects the content control. This causes Word to scroll to the selection.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param selectionMode - Optional. The selection mode must be `select`, `start`, or `end`. `select` is the default.
         */
        select(selectionMode?: "Select" | "Start" | "End"): void;
        
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
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.ContentControl;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.ContentControl;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Word.ContentControl` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ContentControlData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.ContentControlData;
    }
    /**
     * Contains a collection of {@link Word.ContentControl} objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text, plain text, checkbox, dropdown list, and combo box content controls are supported.
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
         * Gets a content control by its identifier. Throws an `ItemNotFound` error if there isn't a content control with the identifier in this collection.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param id - A content control identifier.
         */
        getById(id: number): Word.ContentControl;
        
        /**
         * Gets the content controls that have the specified tag.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param tag - A tag set on a content control.
         */
        getByTag(tag: string): Word.ContentControlCollection;
        /**
         * Gets the content controls that have the specified title.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param title - The title of a content control.
         */
        getByTitle(title: string): Word.ContentControlCollection;
        
        
        
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
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.ContentControlCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.ContentControlCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Word.ContentControlCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ContentControlCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.ContentControlCollectionData;
    }
    
    
    
    
    
    
    
    
    
    /**
     * The `Document` object is the top level object. A `Document` object contains one or more sections, content controls, and the body that contains the contents of the document.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    export class Document extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        
        
        
        /**
         * Gets the `Body` object of the main document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly body: Word.Body;
        
        
        
        
        
        
        /**
         * Gets the collection of `ContentControl` objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly contentControls: Word.ContentControlCollection;
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        /**
         * Gets the collection of `Section` objects in the document.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly sections: Word.SectionCollection;
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        /**
         * Indicates whether the changes in the document have been saved. A value of `true` indicates that the document hasn't changed since it was saved.
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
         * Gets the current selection of the document. Multiple selections aren't supported.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        getSelection(): Word.Range;
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        /**
         * Saves the document.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * Note: The `saveBehavior` and `fileName` parameters were introduced in WordApi 1.5.
         *
         * @param saveBehavior - Optional. The save behavior must be `save` or `prompt`. Default value is `save`.
         * @param fileName - Optional. The file name (exclude file extension). Only takes effect for a new document.
         */
        save(saveBehavior?: Word.SaveBehavior, fileName?: string): void;
        /**
         * Saves the document.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * Note: The `saveBehavior` and `fileName` parameters were introduced in WordApi 1.5.
         *
         * @param saveBehavior - Optional. The save behavior must be `save` or `prompt`. Default value is `save`.
         * @param fileName - Optional. The file name (exclude file extension). Only takes effect for a new document.
         */
        save(saveBehavior?: "Save" | "Prompt", fileName?: string): void;
        
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
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.Document;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.Document;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Word.Document` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.DocumentData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Word.Interfaces.DocumentData;
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
         * Specifies whether the font is bold. `true` if the font is formatted as bold, otherwise, `false`.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        bold: boolean;
        
        /**
         * Specifies the color for the font. You can provide the value in the '#RRGGBB' format or the color name.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        color: string;
        
        
        
        
        
        /**
         * Specifies whether the font has a double strikethrough. `true` if the font is formatted as double strikethrough text, otherwise, `false`.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        doubleStrikeThrough: boolean;
        
        
        
        
        /**
         * Specifies the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to `null`. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or `null` for no highlight color. Note: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        highlightColor: string;
        /**
         * Specifies whether the font is italicized. `true` if the font is italicized, otherwise, `false`.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        italic: boolean;
        
        
        
        /**
         * Specifies the name of the font.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        name: string;
        
        
        
        
        
        
        
        
        
        
        /**
         * Specifies the font size in points.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        size: number;
        
        
        
        /**
         * Specifies whether the font has a strikethrough. `true` if the font is formatted as strikethrough text, otherwise, `false`.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        strikeThrough: boolean;
        
        /**
         * Specifies whether the font is a subscript. `true` if the font is formatted as subscript, otherwise, `false`.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        subscript: boolean;
        /**
         * Specifies whether the font is a superscript. `true` if the font is formatted as superscript, otherwise, `false`.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        superscript: boolean;
        /**
         * Specifies the font's underline type. `none` if the font isn't underlined.
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
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.Font;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.Font;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Word.Font` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.FontData`) that contains shallow copies of any loaded child properties from the original object.
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
         * Gets the parent paragraph that contains the inline image.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         */
        readonly paragraph: Word.Paragraph;
        /**
         * Gets the content control that contains the inline image. Throws an `ItemNotFound` error if there isn't a parent content control.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly parentContentControl: Word.ContentControl;
        
        
        
        
        
        /**
         * Specifies a string that represents the alternative text associated with the inline image.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        altTextDescription: string;
        /**
         * Specifies a string that contains the title for the inline image.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        altTextTitle: string;
        /**
         * Specifies a number that describes the height of the inline image.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        height: number;
        /**
         * Specifies a hyperlink on the image. Use a '#' to separate the address part from the optional location part.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        hyperlink: string;
        
        /**
         * Specifies a value that indicates whether the inline image retains its original proportions when you resize it.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        lockAspectRatio: boolean;
        /**
         * Specifies a number that describes the width of the inline image.
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
         * Gets the Base64-encoded string representation of the inline image.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        getBase64ImageSrc(): OfficeExtension.ClientResult<string>;
        
        
        
        /**
         * Inserts a break at the specified location in the main document.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param breakType - The break type to add.
         * @param insertLocation - The value must be `before` or `after`.
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
         * Insertion isn't supported if the document being inserted contains an ActiveX control (likely in a form field). Consider replacing such a form field with a content control or other option appropriate for your scenario.
         *
         * @param base64File - The Base64-encoded content of a .docx file.
         * @param insertLocation - The value must be `before` or `after`.
         */
        insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Range;
        /**
         * Inserts HTML at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param html - The HTML to be inserted.
         * @param insertLocation - The value must be `before` or `after`.
         */
        insertHtml(html: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Range;
        /**
         * Inserts an inline picture at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - The Base64-encoded image to be inserted.
         * @param insertLocation - The value must be `replace`, `before`, or `after`.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.before | Word.InsertLocation.after | "Replace" | "Before" | "After"): Word.InlinePicture;
        /**
         * Inserts OOXML at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param ooxml - The OOXML to be inserted.
         * @param insertLocation - The value must be `before` or `after`.
         */
        insertOoxml(ooxml: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Range;
        /**
         * Inserts a paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param paragraphText - The paragraph text to be inserted.
         * @param insertLocation - The value must be `before` or `after`.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Paragraph;
        /**
         * Inserts text at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param text - Text to be inserted.
         * @param insertLocation - The value must be `before` or `after`.
         */
        insertText(text: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Range;
        /**
         * Selects the inline picture. This causes Word to scroll to the selection.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param selectionMode - Optional. The selection mode must be `select`, `start`, or `end`. `select` is the default.
         */
        select(selectionMode?: Word.SelectionMode): void;
        /**
         * Selects the inline picture. This causes Word to scroll to the selection.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param selectionMode - Optional. The selection mode must be `select`, `start`, or `end`. `select` is the default.
         */
        select(selectionMode?: "Select" | "Start" | "End"): void;
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
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.InlinePicture;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.InlinePicture;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Word.InlinePicture` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.InlinePictureData`) that contains shallow copies of any loaded child properties from the original object.
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
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.InlinePictureCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.InlinePictureCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Word.InlinePictureCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.InlinePictureCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.InlinePictureCollectionData;
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
         * Gets the collection of `ContentControl` objects in the paragraph.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly contentControls: Word.ContentControlCollection;
        
        
        /**
         * Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly font: Word.Font;
        
        /**
         * Gets the collection of `InlinePicture` objects in the paragraph. The collection doesn't include floating images.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly inlinePictures: Word.InlinePictureCollection;
        
        
        
        
        
        /**
         * Gets the content control that contains the paragraph. Throws an `ItemNotFound` error if there isn't a parent content control.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly parentContentControl: Word.ContentControl;
        
        
        
        
        
        
        
        /**
         * Specifies the alignment for a paragraph. The value can be `left`, `centered`, `right`, or `justified`.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        alignment: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
        /**
         * Specifies the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        firstLineIndent: number;
        
        
        /**
         * Specifies the left indent value, in points, for the paragraph.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        leftIndent: number;
        /**
         * Specifies the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        lineSpacing: number;
        /**
         * Specifies the amount of spacing, in grid lines, after the paragraph.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        lineUnitAfter: number;
        /**
         * Specifies the amount of spacing, in grid lines, before the paragraph.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        lineUnitBefore: number;
        /**
         * Specifies the outline level for the paragraph.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        outlineLevel: number;
        /**
         * Specifies the right indent value, in points, for the paragraph.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        rightIndent: number;
        /**
         * Specifies the spacing, in points, after the paragraph.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        spaceAfter: number;
        /**
         * Specifies the spacing, in points, before the paragraph.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        spaceBefore: number;
        /**
         * Specifies the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        style: string;
        
        
        /**
         * Gets the text of the paragraph.
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
         * Clears the contents of the `Paragraph` object. The user can perform the undo operation on the cleared content.
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
         * Gets an HTML representation of the `Paragraph` object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for the formatting of the document. This method doesn't return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use `Paragraph.getOoxml()` and convert the returned XML to HTML.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        
        
        /**
         * Gets the Office Open XML (OOXML) representation of the `Paragraph` object.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        
        
        
        
        
                
        
        
        
        
        
        /**
         * Inserts a break at the specified location in the main document.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param breakType - The break type to add to the document.
         * @param insertLocation - The value must be `before` or `after`.
         */
        insertBreak(breakType: Word.BreakType | "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): void;
        
        /**
         * Wraps the `Paragraph` object with a content control.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * Note: The `contentControlType` parameter was introduced in WordApi 1.5. `plainText` support was added in WordApi 1.5. `checkBox` support was added in WordApi 1.7.
         * `dropDownList` and `comboBox` support was added in WordApi 1.9. Support for `buildingBlockGallery`, `datePicker`, `group`, `picture`, and `repeatingSection` was added in WordApiDesktop 1.3.
         *
         * @param contentControlType - Optional. Content control type to insert. Must be `richText`, `plainText`, `checkBox`, `dropDownList`, `comboBox`, `buildingBlockGallery`, `datePicker`, `repeatingSection`, `picture`, or `group`. The default is `richText`.
         */
        insertContentControl(contentControlType?: Word.ContentControlType.richText | Word.ContentControlType.plainText | Word.ContentControlType.checkBox | Word.ContentControlType.dropDownList | Word.ContentControlType.comboBox | Word.ContentControlType.buildingBlockGallery | Word.ContentControlType.datePicker | Word.ContentControlType.repeatingSection | Word.ContentControlType.picture | Word.ContentControlType.group | "RichText" | "PlainText" | "CheckBox" | "DropDownList" | "ComboBox" | "BuildingBlockGallery" | "DatePicker" | "RepeatingSection" | "Picture" | "Group"): Word.ContentControl;
        /**
         * Inserts a document into the paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * Insertion isn't supported if the document being inserted contains an ActiveX control (likely in a form field). Consider replacing such a form field with a content control or other option appropriate for your scenario.
         *
         * @param base64File - The Base64-encoded content of a .docx file.
         * @param insertLocation - The value must be `replace`, `start`, or `end`.
         */
        insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        
        
        /**
         * Inserts HTML into the paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param html - The HTML to be inserted in the paragraph.
         * @param insertLocation - The value must be `replace`, `start`, or `end`.
         */
        insertHtml(html: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Inserts a picture into the paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param base64EncodedImage - The Base64-encoded image to be inserted.
         * @param insertLocation - The value must be `replace`, `start`, or `end`.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.InlinePicture;
        /**
         * Inserts OOXML into the paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - The OOXML to be inserted in the paragraph.
         * @param insertLocation - The value must be `replace`, `start`, or `end`.
         */
        insertOoxml(ooxml: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        /**
         * Inserts a paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - The paragraph text to be inserted.
         * @param insertLocation - The value must be `before` or `after`.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Paragraph;
        
        
        /**
         * Inserts text into the paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param text - Text to be inserted.
         * @param insertLocation - The value must be `replace`, `start`, or `end`.
         */
        insertText(text: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
        
        
        
        
        
        
        
        
        
        
        
        
        /**
         * Performs a search with the specified search options on the scope of the `Paragraph` object. The search results are a collection of `Range` objects.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param searchText - The search text.
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
         * @param selectionMode - Optional. The selection mode must be `select`, `start`, or `end`. `select` is the default.
         */
        select(selectionMode?: Word.SelectionMode): void;
        /**
         * Selects and navigates the Word UI to the paragraph.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param selectionMode - Optional. The selection mode must be `select`, `start`, or `end`. `select` is the default.
         */
        select(selectionMode?: "Select" | "Start" | "End"): void;
        
        
        
        
        
        
        
        
        
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
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.Paragraph;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.Paragraph;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Word.Paragraph` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ParagraphData`) that contains shallow copies of any loaded child properties from the original object.
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
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.ParagraphCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.ParagraphCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
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
         * Gets the collection of `ContentControl` objects in the range.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly contentControls: Word.ContentControlCollection;
        
        
        
        /**
         * Gets the text format of the range. Use this to get and set font name, size, color, and other properties.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly font: Word.Font;
        
        
        
        /**
         * Gets the collection of `InlinePicture` objects in the range.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         */
        readonly inlinePictures: Word.InlinePictureCollection;
        
        
        
        /**
         * Gets the collection of `Paragraph` objects in the range.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * Important: For requirement sets 1.1 and 1.2, paragraphs in tables wholly contained within this range aren't returned. From requirement set 1.3, paragraphs in such tables are also returned.
         */
        readonly paragraphs: Word.ParagraphCollection;
        
        /**
         * Gets the currently supported content control that contains the range. Throws an `ItemNotFound` error if there isn't a parent content control.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        readonly parentContentControl: Word.ContentControl;
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        /**
         * Specifies the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        style: string;
        
        /**
         * Gets the text of the range.
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
         * Clears the contents of the `Range` object. The user can perform the undo operation on the cleared content.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        clear(): void;
        
        /**
         * Deletes the range and its content from the document.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        delete(): void;
        
        
        
        
        
        
        /**
         * Gets an HTML representation of the `Range` object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method doesn't return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use `Range.getOoxml()` and convert the returned XML to HTML.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        
        
        
        /**
         * Gets the OOXML representation of the `Range` object.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        getOoxml(): OfficeExtension.ClientResult<string>;
        
        
        
        
        
        
        
        /**
         * Inserts a break at the specified location in the main document.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param breakType - The break type to add.
         * @param insertLocation - The value must be `before` or `after`.
         */
        insertBreak(breakType: Word.BreakType | "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): void;
        
        
        /**
         * Wraps the `Range` object with a content control.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * Note: The `contentControlType` parameter was introduced in WordApi 1.5. `plainText` support was added in WordApi 1.5. `checkBox` support was added in WordApi 1.7.
         * `dropDownList` and `comboBox` support was added in WordApi 1.9. Support for `buildingBlockGallery`, `datePicker`, `group`, `picture`, and `repeatingSection` was added in WordApiDesktop 1.3.
         *
         * @param contentControlType - Optional. Content control type to insert. Must be `richText`, `plainText`, `checkBox`, `dropDownList`, `comboBox`, `buildingBlockGallery`, `datePicker`, `repeatingSection`, `picture`, or `group`. The default is `richText`.
         */
        insertContentControl(contentControlType?: Word.ContentControlType.richText | Word.ContentControlType.plainText | Word.ContentControlType.checkBox | Word.ContentControlType.dropDownList | Word.ContentControlType.comboBox | Word.ContentControlType.buildingBlockGallery | Word.ContentControlType.datePicker | Word.ContentControlType.repeatingSection | Word.ContentControlType.picture | Word.ContentControlType.group | "RichText" | "PlainText" | "CheckBox" | "DropDownList" | "ComboBox" | "BuildingBlockGallery" | "DatePicker" | "RepeatingSection" | "Picture" | "Group"): Word.ContentControl;
        
        
        
        /**
         * Inserts a document at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * Insertion isn't supported if the document being inserted contains an ActiveX control (likely in a form field). Consider replacing such a form field with a content control or other option appropriate for your scenario.
         *
         * @param base64File - The Base64-encoded content of a .docx file.
         * @param insertLocation - The value must be `replace`, `start`, `end`, `before`, or `after`.
         */
        insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation | "Replace" | "Start" | "End" | "Before" | "After"): Word.Range;
        
        
        
        /**
         * Inserts HTML at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param html - The HTML to be inserted.
         * @param insertLocation - The value must be `replace`, `start`, `end`, `before`, or `after`.
         */
        insertHtml(html: string, insertLocation: Word.InsertLocation | "Replace" | "Start" | "End" | "Before" | "After"): Word.Range;
        /**
         * Inserts a picture at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.2]
         *
         * @param base64EncodedImage - The Base64-encoded image to be inserted.
         * @param insertLocation - The value must be `replace`, `start`, `end`, `before`, or `after`.
         */
        insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation | "Replace" | "Start" | "End" | "Before" | "After"): Word.InlinePicture;
        /**
         * Inserts OOXML at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param ooxml - The OOXML to be inserted.
         * @param insertLocation - The value must be `replace`, `start`, `end`, `before`, or `after`.
         */
        insertOoxml(ooxml: string, insertLocation: Word.InsertLocation | "Replace" | "Start" | "End" | "Before" | "After"): Word.Range;
        /**
         * Inserts a paragraph at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param paragraphText - The paragraph text to be inserted.
         * @param insertLocation - The value must be `before` or `after`.
         */
        insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Paragraph;
        
        
        /**
         * Inserts text at the specified location.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param text - Text to be inserted.
         * @param insertLocation - The value must be `replace`, `start`, `end`, `before`, or `after`.
         */
        insertText(text: string, insertLocation: Word.InsertLocation | "Replace" | "Start" | "End" | "Before" | "After"): Word.Range;
        
        
        
        
        /**
         * Performs a search with the specified search options on the scope of the `Range` object. The search results are a collection of `Range` objects.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param searchText - The search text.
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
         * @param selectionMode - Optional. The selection mode must be `select`, `start`, or `end`. `select` is the default.
         */
        select(selectionMode?: Word.SelectionMode): void;
        /**
         * Selects and navigates the Word UI to the range.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param selectionMode - Optional. The selection mode must be `select`, `start`, or `end`. `select` is the default.
         */
        select(selectionMode?: "Select" | "Start" | "End"): void;
        
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
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.Range;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.Range;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Word.Range` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.RangeData`) that contains shallow copies of any loaded child properties from the original object.
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
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.RangeCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.RangeCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Word.RangeCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.RangeCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.RangeCollectionData;
    }
    
    
    /**
     * Specifies the options to be included in a search operation.
                To learn more about how to use search options in the Word JavaScript APIs, read {@link https://learn.microsoft.com/office/dev/add-ins/word/search-option-guidance | Use search options to find text in your Word add-in}.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    export class SearchOptions extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * If provided, specifies whether to ignore all punctuation characters between words. The default is `false`. Corresponds to the _Ignore punctuation characters_ check box in the **Find and Replace** dialog box.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        ignorePunct: boolean;
        /**
         * If provided, specifies whether to ignore all whitespace between words. The default is `false`. Corresponds to the _Ignore white-space characters_ check box in the **Find and Replace** dialog box.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        ignoreSpace: boolean;
        /**
         * If provided, specifies whether to perform a case sensitive search. The default is `false`. Corresponds to the _Match case_ check box in the **Find and Replace** dialog box.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        matchCase: boolean;
        /**
         * If provided, specifies whether to match words that begin with the search string. The default is `false`. Corresponds to the _Match prefix_ check box in the **Find and Replace** dialog box.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        matchPrefix: boolean;
        /**
         * If provided, specifies whether to match words that end with the search string. The default is `false`. Corresponds to the _Match suffix_ check box in the **Find and Replace** dialog box.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        matchSuffix: boolean;
        /**
         * If provided, specifies whether to find only entire words, not text that's part of a larger word. The default is `false`. Corresponds to the _Find whole words only_ check box in the **Find and Replace** dialog box.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        matchWholeWord: boolean;
        /**
         * If provided, specifies whether the search will be performed using special search operators. The default is `false`. Corresponds to the _Use wildcards_ check box in the **Find and Replace** dialog box.
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
         * Create a new instance of the `Word.SearchOptions` object.
         */
        static newObject(context: OfficeExtension.ClientRequestContext): Word.SearchOptions;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Word.SearchOptions` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SearchOptionsData`) that contains shallow copies of any loaded child properties from the original object.
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
         * Gets the `Body` object of the section. This doesn't include the header, footer, and other section metadata.
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
         * @param type - The type of footer to return. This value must be: `primary`, `firstPage`, or `evenPages`.
         */
        getFooter(type: Word.HeaderFooterType): Word.Body;
        /**
         * Gets one of the section's footers.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param type - The type of footer to return. This value must be: `primary`, `firstPage`, or `evenPages`.
         */
        getFooter(type: "Primary" | "FirstPage" | "EvenPages"): Word.Body;
        /**
         * Gets one of the section's headers.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param type - The type of header to return. This value must be: `primary`, `firstPage`, or `evenPages`.
         */
        getHeader(type: Word.HeaderFooterType): Word.Body;
        /**
         * Gets one of the section's headers.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         *
         * @param type - The type of header to return. This value must be: `primary`, `firstPage`, or `evenPages`.
         */
        getHeader(type: "Primary" | "FirstPage" | "EvenPages"): Word.Body;
        
        
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
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.Section;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.Section;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Word.Section` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SectionData`) that contains shallow copies of any loaded child properties from the original object.
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
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.add(thisObject)}. If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
         */
        track(): Word.SectionCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for {@link https://learn.microsoft.com/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member | context.trackedObjects.remove(thisObject)}. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.
         */
        untrack(): Word.SectionCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `Word.SectionCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SectionCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Word.Interfaces.SectionCollectionData;
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    /**
     * Specifies supported content control types and subtypes.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    enum ContentControlType {
        /**
         * Unknown content control type.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        unknown = "Unknown",
        /**
         * Rich text content control subtype containing inline elements.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        richTextInline = "RichTextInline",
        /**
         * Rich text content control subtype containing paragraphs.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        richTextParagraphs = "RichTextParagraphs",
        /**
         * Rich text content control subtype containing a whole cell.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        richTextTableCell = "RichTextTableCell",
        /**
         * Rich text content control subtype containing a whole row.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        richTextTableRow = "RichTextTableRow",
        /**
         * Rich text content control subtype containing a whole table.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        richTextTable = "RichTextTable",
        /**
         * Plain text content control subtype containing inline elements.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        plainTextInline = "PlainTextInline",
        /**
         * Plain text content control subtype containing paragraphs.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        plainTextParagraph = "PlainTextParagraph",
        /**
         * Picture content control (type and subtype).
         * @remarks
         * [Api set: WordApi 1.1]
         */
        picture = "Picture",
        /**
         * Building block gallery content control (type and subtype).
         * @remarks
         * [Api set: WordApi 1.1]
         */
        buildingBlockGallery = "BuildingBlockGallery",
        /**
         * Check box content control (type and subtype).
         * @remarks
         * [Api set: WordApi 1.1]
         */
        checkBox = "CheckBox",
        /**
         * Combo box content control (type and subtype).
         * @remarks
         * [Api set: WordApi 1.1]
         */
        comboBox = "ComboBox",
        /**
         * Dropdown list content control (type and subtype).
         * @remarks
         * [Api set: WordApi 1.1]
         */
        dropDownList = "DropDownList",
        /**
         * Date picker content control (type and subtype).
         * @remarks
         * [Api set: WordApi 1.1]
         */
        datePicker = "DatePicker",
        /**
         * Repeating section content control (type and subtype).
         * @remarks
         * [Api set: WordApi 1.1]
         */
        repeatingSection = "RepeatingSection",
        /**
         * Rich text content control type.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        richText = "RichText",
        /**
         * Plain text content control type.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        plainText = "PlainText",
        /**
         * Group content control type.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        group = "Group",
    }
    /**
     * Represents the content control appearance.
     *
     * @remarks
     * [Api set: WordApi 1.1]
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
         * Represents a content control that isn't shown.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        hidden = "Hidden",
    }
    
    /**
     * Represents the supported styles for underline format.
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
         * Warning: hidden has been deprecated.
         * @deprecated Hidden is no longer supported.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        hidden = "Hidden",
        /**
         * Warning: dotLine has been deprecated.
         * @deprecated DotLine is no longer supported.
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
         * A heavy dotted underline.
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
         * A heavy dash underline.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        dashLineHeavy = "DashLineHeavy",
        /**
         * A long dash underline.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        dashLineLong = "DashLineLong",
        /**
         * A long heavy dash underline.
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
         * A heavy dot-dash underline.
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
         * An alternating heavy dot-dot-dash underline.
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
         * A heavy wavy underline.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        waveHeavy = "WaveHeavy",
        /**
         * A double wavy underline.
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
         * Warning: next has been deprecated. Use sectionNext instead.
         * @deprecated Use sectionNext instead.
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
     * Represents the insertion location type.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     *
     * To be used with an API call, such as `obj.insertSomething(newStuff, location);`.
     * If the location is `before` or `after`, the new content will be outside of the modified object.
     * If the location is `start` or `end`, the new content will be included as part of the modified object.
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
     * Represents the horizontal alignment of text.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    enum Alignment {
        /**
         * Mixed horizontal alignment.
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
     * Represents the type of header or footer.
     *
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
     * Represents where the cursor (insertion point) is set in the document after a selection.
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
     * Specifies the save behavior for `Document.save`.
     *
     * @remarks
     * [Api set: WordApi 1.1]
     */
    enum SaveBehavior {
        /**
         * Saves the document without prompting the user. If it's a new document,
                    it will be saved with the default name or specified name in the default location.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        save = "Save",
        /**
         * Displays the "Save As" dialog to the user if the document hasn't been saved.
                    Won't take effect if the document was previously saved.
         * @remarks
         * [Api set: WordApi 1.1]
         */
        prompt = "Prompt",
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    enum ErrorCodes {
        accessDenied = "AccessDenied",
        generalException = "GeneralException",
        invalidArgument = "InvalidArgument",
        itemNotFound = "ItemNotFound",
        notAllowed = "NotAllowed",
        notImplemented = "NotImplemented",
        searchDialogIsOpen = "SearchDialogIsOpen",
        searchStringInvalidOrTooLong = "SearchStringInvalidOrTooLong",
    }
    export namespace Interfaces {
        /**
        * Provides ways to load properties of only a subset of members of a collection.
        */
        export interface CollectionLoadOptions {
            /**
            * Specify the number of items in the queried collection to be included in the result.
            */
            $top?: number;
            /**
            * Specify the number of items in the collection that are to be skipped and not included in the result. If `top` is specified, the selection of result will start after skipping the specified number of items.
            */
            $skip?: number;
        }
        /** An interface for updating data on the `Editor` object, for use in `editor.set({ ... })`. */
        export interface EditorUpdateData {
            
            
        }
        /** An interface for updating data on the `ConflictCollection` object, for use in `conflictCollection.set({ ... })`. */
        export interface ConflictCollectionUpdateData {
            items?: Word.Interfaces.ConflictData[];
        }
        /** An interface for updating data on the `Conflict` object, for use in `conflict.set({ ... })`. */
        export interface ConflictUpdateData {
            
        }
        /** An interface for updating data on the `AnnotationCollection` object, for use in `annotationCollection.set({ ... })`. */
        export interface AnnotationCollectionUpdateData {
            items?: Word.Interfaces.AnnotationData[];
        }
        /** An interface for updating data on the `Application` object, for use in `application.set({ ... })`. */
        export interface ApplicationUpdateData {
            
            
        }
        /** An interface for updating data on the `Body` object, for use in `body.set({ ... })`. */
        export interface BodyUpdateData {
            /**
            * Gets the text format of the body. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontUpdateData;
            /**
             * Specifies the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: string;
            
        }
        /** An interface for updating data on the `Border` object, for use in `border.set({ ... })`. */
        export interface BorderUpdateData {
            
            
            
            
        }
        /** An interface for updating data on the `BorderUniversal` object, for use in `borderUniversal.set({ ... })`. */
        export interface BorderUniversalUpdateData {
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `BorderCollection` object, for use in `borderCollection.set({ ... })`. */
        export interface BorderCollectionUpdateData {
            
            
            
            
            
            
            items?: Word.Interfaces.BorderData[];
        }
        /** An interface for updating data on the `BorderUniversalCollection` object, for use in `borderUniversalCollection.set({ ... })`. */
        export interface BorderUniversalCollectionUpdateData {
            items?: Word.Interfaces.BorderUniversalData[];
        }
        /** An interface for updating data on the `Break` object, for use in `break.set({ ... })`. */
        export interface BreakUpdateData {
            
        }
        /** An interface for updating data on the `BreakCollection` object, for use in `breakCollection.set({ ... })`. */
        export interface BreakCollectionUpdateData {
            items?: Word.Interfaces.BreakData[];
        }
        /** An interface for updating data on the `BuildingBlock` object, for use in `buildingBlock.set({ ... })`. */
        export interface BuildingBlockUpdateData {
            
            
            
            
        }
        /** An interface for updating data on the `CheckboxContentControl` object, for use in `checkboxContentControl.set({ ... })`. */
        export interface CheckboxContentControlUpdateData {
            
        }
        /** An interface for updating data on the `CoauthoringLock` object, for use in `coauthoringLock.set({ ... })`. */
        export interface CoauthoringLockUpdateData {
            
        }
        /** An interface for updating data on the `CoauthoringLockCollection` object, for use in `coauthoringLockCollection.set({ ... })`. */
        export interface CoauthoringLockCollectionUpdateData {
            items?: Word.Interfaces.CoauthoringLockData[];
        }
        /** An interface for updating data on the `CoauthorCollection` object, for use in `coauthorCollection.set({ ... })`. */
        export interface CoauthorCollectionUpdateData {
            items?: Word.Interfaces.CoauthorData[];
        }
        /** An interface for updating data on the `CoauthoringUpdate` object, for use in `coauthoringUpdate.set({ ... })`. */
        export interface CoauthoringUpdateUpdateData {
            
        }
        /** An interface for updating data on the `CoauthoringUpdateCollection` object, for use in `coauthoringUpdateCollection.set({ ... })`. */
        export interface CoauthoringUpdateCollectionUpdateData {
            items?: Word.Interfaces.CoauthoringUpdateData[];
        }
        /** An interface for updating data on the `Comment` object, for use in `comment.set({ ... })`. */
        export interface CommentUpdateData {
            
            
            
        }
        /** An interface for updating data on the `CommentCollection` object, for use in `commentCollection.set({ ... })`. */
        export interface CommentCollectionUpdateData {
            items?: Word.Interfaces.CommentData[];
        }
        /** An interface for updating data on the `CommentContentRange` object, for use in `commentContentRange.set({ ... })`. */
        export interface CommentContentRangeUpdateData {
            
            
            
            
            
        }
        /** An interface for updating data on the `CommentReply` object, for use in `commentReply.set({ ... })`. */
        export interface CommentReplyUpdateData {
            
            
            
        }
        /** An interface for updating data on the `CommentReplyCollection` object, for use in `commentReplyCollection.set({ ... })`. */
        export interface CommentReplyCollectionUpdateData {
            items?: Word.Interfaces.CommentReplyData[];
        }
        /** An interface for updating data on the `ConditionalStyle` object, for use in `conditionalStyle.set({ ... })`. */
        export interface ConditionalStyleUpdateData {
            
            
            
            
        }
        /** An interface for updating data on the `XmlMapping` object, for use in `xmlMapping.set({ ... })`. */
        export interface XmlMappingUpdateData {
            
            
        }
        /** An interface for updating data on the `CustomXmlPrefixMappingCollection` object, for use in `customXmlPrefixMappingCollection.set({ ... })`. */
        export interface CustomXmlPrefixMappingCollectionUpdateData {
            items?: Word.Interfaces.CustomXmlPrefixMappingData[];
        }
        /** An interface for updating data on the `CustomXmlSchemaCollection` object, for use in `customXmlSchemaCollection.set({ ... })`. */
        export interface CustomXmlSchemaCollectionUpdateData {
            items?: Word.Interfaces.CustomXmlSchemaData[];
        }
        /** An interface for updating data on the `CustomXmlNodeCollection` object, for use in `customXmlNodeCollection.set({ ... })`. */
        export interface CustomXmlNodeCollectionUpdateData {
            items?: Word.Interfaces.CustomXmlNodeData[];
        }
        /** An interface for updating data on the `CustomXmlNode` object, for use in `customXmlNode.set({ ... })`. */
        export interface CustomXmlNodeUpdateData {
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `ContentControl` object, for use in `contentControl.set({ ... })`. */
        export interface ContentControlUpdateData {
            
            
            
            /**
            * Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontUpdateData;
            
            
            
            
            /**
             * Specifies the appearance of the content control. The value can be `boundingBox`, `tags`, or `hidden`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            appearance?: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden";
            /**
             * Specifies a value that indicates whether the user can delete the content control. Mutually exclusive with `removeWhenEdited`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            cannotDelete?: boolean;
            /**
             * Specifies a value that indicates whether the user can edit the contents of the content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            cannotEdit?: boolean;
            /**
             * Specifies the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            color?: string;
            /**
             * Specifies the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            placeholderText?: string;
            /**
             * Specifies a value that indicates whether the content control is removed after it's edited. Mutually exclusive with `cannotDelete`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            removeWhenEdited?: boolean;
            /**
             * Specifies the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the `styleBuiltIn` property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: string;
            
            /**
             * Specifies a tag to identify a content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            tag?: string;
            /**
             * Specifies the title for a content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            title?: string;
        }
        /** An interface for updating data on the `ContentControlCollection` object, for use in `contentControlCollection.set({ ... })`. */
        export interface ContentControlCollectionUpdateData {
            items?: Word.Interfaces.ContentControlData[];
        }
        /** An interface for updating data on the `ContentControlListItem` object, for use in `contentControlListItem.set({ ... })`. */
        export interface ContentControlListItemUpdateData {
            
            
            
        }
        /** An interface for updating data on the `ContentControlListItemCollection` object, for use in `contentControlListItemCollection.set({ ... })`. */
        export interface ContentControlListItemCollectionUpdateData {
            items?: Word.Interfaces.ContentControlListItemData[];
        }
        /** An interface for updating data on the `CustomProperty` object, for use in `customProperty.set({ ... })`. */
        export interface CustomPropertyUpdateData {
            
        }
        /** An interface for updating data on the `CustomPropertyCollection` object, for use in `customPropertyCollection.set({ ... })`. */
        export interface CustomPropertyCollectionUpdateData {
            items?: Word.Interfaces.CustomPropertyData[];
        }
        /** An interface for updating data on the `CustomXmlPart` object, for use in `customXmlPart.set({ ... })`. */
        export interface CustomXmlPartUpdateData {
            
        }
        /** An interface for updating data on the `CustomXmlPartCollection` object, for use in `customXmlPartCollection.set({ ... })`. */
        export interface CustomXmlPartCollectionUpdateData {
            items?: Word.Interfaces.CustomXmlPartData[];
        }
        /** An interface for updating data on the `CustomXmlPartScopedCollection` object, for use in `customXmlPartScopedCollection.set({ ... })`. */
        export interface CustomXmlPartScopedCollectionUpdateData {
            items?: Word.Interfaces.CustomXmlPartData[];
        }
        /** An interface for updating data on the `Document` object, for use in `document.set({ ... })`. */
        export interface DocumentUpdateData {
            
            
            /**
            * Gets the `Body` object of the main document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            body?: Word.Interfaces.BodyUpdateData;
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `DocumentCreated` object, for use in `documentCreated.set({ ... })`. */
        export interface DocumentCreatedUpdateData {
            
            
        }
        /** An interface for updating data on the `DocumentProperties` object, for use in `documentProperties.set({ ... })`. */
        export interface DocumentPropertiesUpdateData {
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `Field` object, for use in `field.set({ ... })`. */
        export interface FieldUpdateData {
            
            
            
            
            
        }
        /** An interface for updating data on the `FieldCollection` object, for use in `fieldCollection.set({ ... })`. */
        export interface FieldCollectionUpdateData {
            items?: Word.Interfaces.FieldData[];
        }
        /** An interface for updating data on the `Font` object, for use in `font.set({ ... })`. */
        export interface FontUpdateData {
            
            
            
            
            
            
            
            
            /**
             * Specifies whether the font is bold. `true` if the font is formatted as bold, otherwise, `false`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            bold?: boolean;
            
            /**
             * Specifies the color for the font. You can provide the value in the '#RRGGBB' format or the color name.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            color?: string;
            
            
            
            
            
            /**
             * Specifies whether the font has a double strikethrough. `true` if the font is formatted as double strikethrough text, otherwise, `false`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            doubleStrikeThrough?: boolean;
            
            
            
            
            /**
             * Specifies the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to `null`. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or `null` for no highlight color. Note: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            highlightColor?: string;
            /**
             * Specifies whether the font is italicized. `true` if the font is italicized, otherwise, `false`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            italic?: boolean;
            
            
            
            /**
             * Specifies the name of the font.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            name?: string;
            
            
            
            
            
            
            
            
            
            
            /**
             * Specifies the font size in points.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            size?: number;
            
            
            
            /**
             * Specifies whether the font has a strikethrough. `true` if the font is formatted as strikethrough text, otherwise, `false`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            strikeThrough?: boolean;
            
            /**
             * Specifies whether the font is a subscript. `true` if the font is formatted as subscript, otherwise, `false`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            subscript?: boolean;
            /**
             * Specifies whether the font is a superscript. `true` if the font is formatted as superscript, otherwise, `false`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            superscript?: boolean;
            /**
             * Specifies the font's underline type. `none` if the font isn't underlined.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            underline?: Word.UnderlineType | "Mixed" | "None" | "Hidden" | "DotLine" | "Single" | "Word" | "Double" | "Thick" | "Dotted" | "DottedHeavy" | "DashLine" | "DashLineHeavy" | "DashLineLong" | "DashLineLongHeavy" | "DotDashLine" | "DotDashLineHeavy" | "TwoDotDashLine" | "TwoDotDashLineHeavy" | "Wave" | "WaveHeavy" | "WaveDouble";
            
        }
        /** An interface for updating data on the `HeadingStyle` object, for use in `headingStyle.set({ ... })`. */
        export interface HeadingStyleUpdateData {
            
            
        }
        /** An interface for updating data on the `HeadingStyleCollection` object, for use in `headingStyleCollection.set({ ... })`. */
        export interface HeadingStyleCollectionUpdateData {
            items?: Word.Interfaces.HeadingStyleData[];
        }
        /** An interface for updating data on the `Hyperlink` object, for use in `hyperlink.set({ ... })`. */
        export interface HyperlinkUpdateData {
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `HyperlinkCollection` object, for use in `hyperlinkCollection.set({ ... })`. */
        export interface HyperlinkCollectionUpdateData {
            items?: Word.Interfaces.HyperlinkData[];
        }
        /** An interface for updating data on the `InlinePicture` object, for use in `inlinePicture.set({ ... })`. */
        export interface InlinePictureUpdateData {
            /**
             * Specifies a string that represents the alternative text associated with the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            altTextDescription?: string;
            /**
             * Specifies a string that contains the title for the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            altTextTitle?: string;
            /**
             * Specifies a number that describes the height of the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            height?: number;
            /**
             * Specifies a hyperlink on the image. Use a '#' to separate the address part from the optional location part.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            hyperlink?: string;
            /**
             * Specifies a value that indicates whether the inline image retains its original proportions when you resize it.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lockAspectRatio?: boolean;
            /**
             * Specifies a number that describes the width of the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            width?: number;
        }
        /** An interface for updating data on the `InlinePictureCollection` object, for use in `inlinePictureCollection.set({ ... })`. */
        export interface InlinePictureCollectionUpdateData {
            items?: Word.Interfaces.InlinePictureData[];
        }
        /** An interface for updating data on the `LinkFormat` object, for use in `linkFormat.set({ ... })`. */
        export interface LinkFormatUpdateData {
            
            
            
            
        }
        /** An interface for updating data on the `ListCollection` object, for use in `listCollection.set({ ... })`. */
        export interface ListCollectionUpdateData {
            items?: Word.Interfaces.ListData[];
        }
        /** An interface for updating data on the `ListItem` object, for use in `listItem.set({ ... })`. */
        export interface ListItemUpdateData {
            
        }
        /** An interface for updating data on the `ListLevel` object, for use in `listLevel.set({ ... })`. */
        export interface ListLevelUpdateData {
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `ListLevelCollection` object, for use in `listLevelCollection.set({ ... })`. */
        export interface ListLevelCollectionUpdateData {
            items?: Word.Interfaces.ListLevelData[];
        }
        /** An interface for updating data on the `ListTemplate` object, for use in `listTemplate.set({ ... })`. */
        export interface ListTemplateUpdateData {
            
            
        }
        /** An interface for updating data on the `NoteItem` object, for use in `noteItem.set({ ... })`. */
        export interface NoteItemUpdateData {
            
            
        }
        /** An interface for updating data on the `NoteItemCollection` object, for use in `noteItemCollection.set({ ... })`. */
        export interface NoteItemCollectionUpdateData {
            items?: Word.Interfaces.NoteItemData[];
        }
        /** An interface for updating data on the `OleFormat` object, for use in `oleFormat.set({ ... })`. */
        export interface OleFormatUpdateData {
            
        /** An interface for updating data on the `PageCollection` object, for use in `pageCollection.set({ ... })`. */
        export interface PageCollectionUpdateData {
            items?: Word.Interfaces.PageData[];
        }
        /** An interface for updating data on the `PaneCollection` object, for use in `paneCollection.set({ ... })`. */
        export interface PaneCollectionUpdateData {
            items?: Word.Interfaces.PaneData[];
        }
        /** An interface for updating data on the `Window` object, for use in `window.set({ ... })`. */
        export interface WindowUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `WindowCollection` object, for use in `windowCollection.set({ ... })`. */
        export interface WindowCollectionUpdateData {
            items?: Word.Interfaces.WindowData[];
        }
        /** An interface for updating data on the `Paragraph` object, for use in `paragraph.set({ ... })`. */
        export interface ParagraphUpdateData {
            /**
            * Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontUpdateData;
            
            
            
            /**
             * Specifies the alignment for a paragraph. The value can be `left`, `centered`, `right`, or `justified`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            alignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
            /**
             * Specifies the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            firstLineIndent?: number;
            /**
             * Specifies the left indent value, in points, for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            leftIndent?: number;
            /**
             * Specifies the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineSpacing?: number;
            /**
             * Specifies the amount of spacing, in grid lines, after the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineUnitAfter?: number;
            /**
             * Specifies the amount of spacing, in grid lines, before the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineUnitBefore?: number;
            /**
             * Specifies the outline level for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            outlineLevel?: number;
            /**
             * Specifies the right indent value, in points, for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            rightIndent?: number;
            /**
             * Specifies the spacing, in points, after the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            spaceAfter?: number;
            /**
             * Specifies the spacing, in points, before the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            spaceBefore?: number;
            /**
             * Specifies the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: string;
            
        }
        /** An interface for updating data on the `ParagraphCollection` object, for use in `paragraphCollection.set({ ... })`. */
        export interface ParagraphCollectionUpdateData {
            items?: Word.Interfaces.ParagraphData[];
        }
        /** An interface for updating data on the `ParagraphFormat` object, for use in `paragraphFormat.set({ ... })`. */
        export interface ParagraphFormatUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `Range` object, for use in `range.set({ ... })`. */
        export interface RangeUpdateData {
            /**
            * Gets the text format of the range. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontUpdateData;
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            /**
             * Specifies the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: string;
            
            
        }
        /** An interface for updating data on the `RangeCollection` object, for use in `rangeCollection.set({ ... })`. */
        export interface RangeCollectionUpdateData {
            items?: Word.Interfaces.RangeData[];
        }
        /** An interface for updating data on the `SearchOptions` object, for use in `searchOptions.set({ ... })`. */
        export interface SearchOptionsUpdateData {
            /**
             * If provided, specifies whether to ignore all punctuation characters between words. The default is `false`. Corresponds to the _Ignore punctuation characters_ check box in the **Find and Replace** dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            ignorePunct?: boolean;
            /**
             * If provided, specifies whether to ignore all whitespace between words. The default is `false`. Corresponds to the _Ignore white-space characters_ check box in the **Find and Replace** dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            ignoreSpace?: boolean;
            /**
             * If provided, specifies whether to perform a case sensitive search. The default is `false`. Corresponds to the _Match case_ check box in the **Find and Replace** dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchCase?: boolean;
            /**
             * If provided, specifies whether to match words that begin with the search string. The default is `false`. Corresponds to the _Match prefix_ check box in the **Find and Replace** dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchPrefix?: boolean;
            /**
             * If provided, specifies whether to match words that end with the search string. The default is `false`. Corresponds to the _Match suffix_ check box in the **Find and Replace** dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchSuffix?: boolean;
            /**
             * If provided, specifies whether to find only entire words, not text that's part of a larger word. The default is `false`. Corresponds to the _Find whole words only_ check box in the **Find and Replace** dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchWholeWord?: boolean;
            /**
             * If provided, specifies whether the search will be performed using special search operators. The default is `false`. Corresponds to the _Use wildcards_ check box in the **Find and Replace** dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchWildcards?: boolean;
        }
        /** An interface for updating data on the `Section` object, for use in `section.set({ ... })`. */
        export interface SectionUpdateData {
            /**
            * Gets the `Body` object of the section. This doesn't include the header, footer, and other section metadata.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            body?: Word.Interfaces.BodyUpdateData;
            
            
        }
        /** An interface for updating data on the `SectionCollection` object, for use in `sectionCollection.set({ ... })`. */
        export interface SectionCollectionUpdateData {
            items?: Word.Interfaces.SectionData[];
        }
        /** An interface for updating data on the `Setting` object, for use in `setting.set({ ... })`. */
        export interface SettingUpdateData {
            
        }
        /** An interface for updating data on the `SettingCollection` object, for use in `settingCollection.set({ ... })`. */
        export interface SettingCollectionUpdateData {
            items?: Word.Interfaces.SettingData[];
        }
        /** An interface for updating data on the `StyleCollection` object, for use in `styleCollection.set({ ... })`. */
        export interface StyleCollectionUpdateData {
            items?: Word.Interfaces.StyleData[];
        }
        /** An interface for updating data on the `Style` object, for use in `style.set({ ... })`. */
        export interface StyleUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `Shading` object, for use in `shading.set({ ... })`. */
        export interface ShadingUpdateData {
            
            
            
        }
        /** An interface for updating data on the `ShadingUniversal` object, for use in `shadingUniversal.set({ ... })`. */
        export interface ShadingUniversalUpdateData {
            
            
            
            
            
        }
        /** An interface for updating data on the `Table` object, for use in `table.set({ ... })`. */
        export interface TableUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `TableStyle` object, for use in `tableStyle.set({ ... })`. */
        export interface TableStyleUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `TabStopCollection` object, for use in `tabStopCollection.set({ ... })`. */
        export interface TabStopCollectionUpdateData {
            items?: Word.Interfaces.TabStopData[];
        }
        /** An interface for updating data on the `TableCollection` object, for use in `tableCollection.set({ ... })`. */
        export interface TableCollectionUpdateData {
            items?: Word.Interfaces.TableData[];
        }
        /** An interface for updating data on the `TableColumn` object, for use in `tableColumn.set({ ... })`. */
        export interface TableColumnUpdateData {
            
            
            
            
        }
        /** An interface for updating data on the `TableColumnCollection` object, for use in `tableColumnCollection.set({ ... })`. */
        export interface TableColumnCollectionUpdateData {
            items?: Word.Interfaces.TableColumnData[];
        }
        /** An interface for updating data on the `TableOfAuthorities` object, for use in `tableOfAuthorities.set({ ... })`. */
        export interface TableOfAuthoritiesUpdateData {
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `TableOfAuthoritiesCollection` object, for use in `tableOfAuthoritiesCollection.set({ ... })`. */
        export interface TableOfAuthoritiesCollectionUpdateData {
            items?: Word.Interfaces.TableOfAuthoritiesData[];
        }
        /** An interface for updating data on the `TableOfAuthoritiesCategoryCollection` object, for use in `tableOfAuthoritiesCategoryCollection.set({ ... })`. */
        export interface TableOfAuthoritiesCategoryCollectionUpdateData {
            items?: Word.Interfaces.TableOfAuthoritiesCategoryData[];
        }
        /** An interface for updating data on the `TableOfContents` object, for use in `tableOfContents.set({ ... })`. */
        export interface TableOfContentsUpdateData {
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `TableOfContentsCollection` object, for use in `tableOfContentsCollection.set({ ... })`. */
        export interface TableOfContentsCollectionUpdateData {
            items?: Word.Interfaces.TableOfContentsData[];
        }
        /** An interface for updating data on the `TableOfFigures` object, for use in `tableOfFigures.set({ ... })`. */
        export interface TableOfFiguresUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `TableOfFiguresCollection` object, for use in `tableOfFiguresCollection.set({ ... })`. */
        export interface TableOfFiguresCollectionUpdateData {
            items?: Word.Interfaces.TableOfFiguresData[];
        }
        /** An interface for updating data on the `TableRow` object, for use in `tableRow.set({ ... })`. */
        export interface TableRowUpdateData {
            
            
            
            
            
            
        }
        /** An interface for updating data on the `TableRowCollection` object, for use in `tableRowCollection.set({ ... })`. */
        export interface TableRowCollectionUpdateData {
            items?: Word.Interfaces.TableRowData[];
        }
        /** An interface for updating data on the `TableCell` object, for use in `tableCell.set({ ... })`. */
        export interface TableCellUpdateData {
            
            
            
            
            
            
        }
        /** An interface for updating data on the `TableCellCollection` object, for use in `tableCellCollection.set({ ... })`. */
        export interface TableCellCollectionUpdateData {
            items?: Word.Interfaces.TableCellData[];
        }
        /** An interface for updating data on the `TableBorder` object, for use in `tableBorder.set({ ... })`. */
        export interface TableBorderUpdateData {
            
            
            
        }
        /** An interface for updating data on the `Template` object, for use in `template.set({ ... })`. */
        export interface TemplateUpdateData {
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `TemplateCollection` object, for use in `templateCollection.set({ ... })`. */
        export interface TemplateCollectionUpdateData {
            items?: Word.Interfaces.TemplateData[];
        }
        /** An interface for updating data on the `TrackedChangeCollection` object, for use in `trackedChangeCollection.set({ ... })`. */
        export interface TrackedChangeCollectionUpdateData {
            items?: Word.Interfaces.TrackedChangeData[];
        }
        /** An interface for updating data on the `View` object, for use in `view.set({ ... })`. */
        export interface ViewUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `Shape` object, for use in `shape.set({ ... })`. */
        export interface ShapeUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `ShapeGroup` object, for use in `shapeGroup.set({ ... })`. */
        export interface ShapeGroupUpdateData {
            
        }
        /** An interface for updating data on the `Canvas` object, for use in `canvas.set({ ... })`. */
        export interface CanvasUpdateData {
            
        }
        /** An interface for updating data on the `ShapeCollection` object, for use in `shapeCollection.set({ ... })`. */
        export interface ShapeCollectionUpdateData {
            items?: Word.Interfaces.ShapeData[];
        }
        /** An interface for updating data on the `ShapeFill` object, for use in `shapeFill.set({ ... })`. */
        export interface ShapeFillUpdateData {
            
            
            
        }
        /** An interface for updating data on the `TextFrame` object, for use in `textFrame.set({ ... })`. */
        export interface TextFrameUpdateData {
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `ShapeTextWrap` object, for use in `shapeTextWrap.set({ ... })`. */
        export interface ShapeTextWrapUpdateData {
            
            
            
            
            
            
        }
        /** An interface for updating data on the `Reviewer` object, for use in `reviewer.set({ ... })`. */
        export interface ReviewerUpdateData {
            
        }
        /** An interface for updating data on the `ReviewerCollection` object, for use in `reviewerCollection.set({ ... })`. */
        export interface ReviewerCollectionUpdateData {
            items?: Word.Interfaces.ReviewerData[];
        }
        /** An interface for updating data on the `RevisionsFilter` object, for use in `revisionsFilter.set({ ... })`. */
        export interface RevisionsFilterUpdateData {
            
            
        }
        /** An interface for updating data on the `RepeatingSectionItem` object, for use in `repeatingSectionItem.set({ ... })`. */
        export interface RepeatingSectionItemUpdateData {
            
        }
        /** An interface for updating data on the `Revision` object, for use in `revision.set({ ... })`. */
        export interface RevisionUpdateData {
            
            
        }
        /** An interface for updating data on the `RevisionCollection` object, for use in `revisionCollection.set({ ... })`. */
        export interface RevisionCollectionUpdateData {
            items?: Word.Interfaces.RevisionData[];
        }
        /** An interface for updating data on the `DatePickerContentControl` object, for use in `datePickerContentControl.set({ ... })`. */
        export interface DatePickerContentControlUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `PictureContentControl` object, for use in `pictureContentControl.set({ ... })`. */
        export interface PictureContentControlUpdateData {
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `GroupContentControl` object, for use in `groupContentControl.set({ ... })`. */
        export interface GroupContentControlUpdateData {
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `BuildingBlockGalleryContentControl` object, for use in `buildingBlockGalleryContentControl.set({ ... })`. */
        export interface BuildingBlockGalleryContentControlUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `RepeatingSectionContentControl` object, for use in `repeatingSectionContentControl.set({ ... })`. */
        export interface RepeatingSectionContentControlUpdateData {
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `ReadabilityStatisticCollection` object, for use in `readabilityStatisticCollection.set({ ... })`. */
        export interface ReadabilityStatisticCollectionUpdateData {
            items?: Word.Interfaces.ReadabilityStatisticData[];
        }
        /** An interface for updating data on the `WebSettings` object, for use in `webSettings.set({ ... })`. */
        export interface WebSettingsUpdateData {
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `XmlNodeCollection` object, for use in `xmlNodeCollection.set({ ... })`. */
        export interface XmlNodeCollectionUpdateData {
            items?: Word.Interfaces.XmlNodeData[];
        }
        /** An interface for updating data on the `XmlNode` object, for use in `xmlNode.set({ ... })`. */
        export interface XmlNodeUpdateData {
            
            
        }
        /** An interface for updating data on the `HtmlDivision` object, for use in `htmlDivision.set({ ... })`. */
        export interface HtmlDivisionUpdateData {
            
            
            
            
        }
        /** An interface for updating data on the `HtmlDivisionCollection` object, for use in `htmlDivisionCollection.set({ ... })`. */
        export interface HtmlDivisionCollectionUpdateData {
            items?: Word.Interfaces.HtmlDivisionData[];
        }
        /** An interface for updating data on the `Frame` object, for use in `frame.set({ ... })`. */
        export interface FrameUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `FrameCollection` object, for use in `frameCollection.set({ ... })`. */
        export interface FrameCollectionUpdateData {
            items?: Word.Interfaces.FrameData[];
        }
        /** An interface for updating data on the `DocumentLibraryVersionCollection` object, for use in `documentLibraryVersionCollection.set({ ... })`. */
        export interface DocumentLibraryVersionCollectionUpdateData {
            items?: Word.Interfaces.DocumentLibraryVersionData[];
        }
        /** An interface for updating data on the `ListFormat` object, for use in `listFormat.set({ ... })`. */
        export interface ListFormatUpdateData {
            
            
        }
        /** An interface for updating data on the `FillFormat` object, for use in `fillFormat.set({ ... })`. */
        export interface FillFormatUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `GlowFormat` object, for use in `glowFormat.set({ ... })`. */
        export interface GlowFormatUpdateData {
            
            
            
        }
        /** An interface for updating data on the `LineFormat` object, for use in `lineFormat.set({ ... })`. */
        export interface LineFormatUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `ReflectionFormat` object, for use in `reflectionFormat.set({ ... })`. */
        export interface ReflectionFormatUpdateData {
            
            
            
            
            
        }
        /** An interface for updating data on the `ColorFormat` object, for use in `colorFormat.set({ ... })`. */
        export interface ColorFormatUpdateData {
            
            
            
            
        }
        /** An interface for updating data on the `ShadowFormat` object, for use in `shadowFormat.set({ ... })`. */
        export interface ShadowFormatUpdateData {
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `ThreeDimensionalFormat` object, for use in `threeDimensionalFormat.set({ ... })`. */
        export interface ThreeDimensionalFormatUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `Bibliography` object, for use in `bibliography.set({ ... })`. */
        export interface BibliographyUpdateData {
            
        }
        /** An interface for updating data on the `SourceCollection` object, for use in `sourceCollection.set({ ... })`. */
        export interface SourceCollectionUpdateData {
            items?: Word.Interfaces.SourceData[];
        }
        /** An interface for updating data on the `PageSetup` object, for use in `pageSetup.set({ ... })`. */
        export interface PageSetupUpdateData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `LineNumbering` object, for use in `lineNumbering.set({ ... })`. */
        export interface LineNumberingUpdateData {
            
            
            
            
            
        }
        /** An interface for updating data on the `TextColumnCollection` object, for use in `textColumnCollection.set({ ... })`. */
        export interface TextColumnCollectionUpdateData {
            items?: Word.Interfaces.TextColumnData[];
        }
        /** An interface for updating data on the `TextColumn` object, for use in `textColumn.set({ ... })`. */
        export interface TextColumnUpdateData {
            
            
        }
        /** An interface for updating data on the `Selection` object, for use in `selection.set({ ... })`. */
        export interface SelectionUpdateData {
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `RangeScopedCollection` object, for use in `rangeScopedCollection.set({ ... })`. */
        export interface RangeScopedCollectionUpdateData {
            items?: Word.Interfaces.RangeData[];
        }
        /** An interface for updating data on the `Bookmark` object, for use in `bookmark.set({ ... })`. */
        export interface BookmarkUpdateData {
            
            
            
        }
        /** An interface for updating data on the `BookmarkCollection` object, for use in `bookmarkCollection.set({ ... })`. */
        export interface BookmarkCollectionUpdateData {
            items?: Word.Interfaces.BookmarkData[];
        }
        /** An interface for updating data on the `Index` object, for use in `index.set({ ... })`. */
        export interface IndexUpdateData {
            
            
        }
        /** An interface for updating data on the `IndexCollection` object, for use in `indexCollection.set({ ... })`. */
        export interface IndexCollectionUpdateData {
            items?: Word.Interfaces.IndexData[];
        }
        /** An interface for updating data on the `ListTemplateCollection` object, for use in `listTemplateCollection.set({ ... })`. */
        export interface ListTemplateCollectionUpdateData {
            items?: Word.Interfaces.ListTemplateData[];
        }
        /** An interface for updating data on the `ListTemplateGalleryCollection` object, for use in `listTemplateGalleryCollection.set({ ... })`. */
        export interface ListTemplateGalleryCollectionUpdateData {
            items?: Word.Interfaces.ListTemplateGalleryData[];
        }
        /** An interface describing the data returned by calling `editor.toJSON()`. */
        export interface EditorData {
            
            
            
            
        }
        /** An interface describing the data returned by calling `conflictCollection.toJSON()`. */
        export interface ConflictCollectionData {
            items?: Word.Interfaces.ConflictData[];
        }
        /** An interface describing the data returned by calling `conflict.toJSON()`. */
        export interface ConflictData {
            
            
        }
        /** An interface describing the data returned by calling `critiqueAnnotation.toJSON()`. */
        export interface CritiqueAnnotationData {
            
        }
        /** An interface describing the data returned by calling `annotation.toJSON()`. */
        export interface AnnotationData {
            
            
        }
        /** An interface describing the data returned by calling `annotationCollection.toJSON()`. */
        export interface AnnotationCollectionData {
            items?: Word.Interfaces.AnnotationData[];
        }
        /** An interface describing the data returned by calling `application.toJSON()`. */
        export interface ApplicationData {
            
            
            
            
        }
        /** An interface describing the data returned by calling `body.toJSON()`. */
        export interface BodyData {
            /**
            * Gets the collection of rich text content control objects in the body.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            contentControls?: Word.Interfaces.ContentControlData[];
            
            /**
            * Gets the text format of the body. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontData;
            /**
            * Gets the collection of `InlinePicture` objects in the body. The collection doesn't include floating images.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            inlinePictures?: Word.Interfaces.InlinePictureData[];
            
            /**
            * Gets the collection of `Paragraph` objects in the body.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            *
            * Important: Paragraphs in tables aren't returned for requirement sets 1.1 and 1.2. From requirement set 1.3, paragraphs in tables are also returned.
            */
            paragraphs?: Word.Interfaces.ParagraphData[];
            
            
            /**
             * Specifies the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: string;
            
            /**
             * Gets the text of the body. Use the insertText method to insert text.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            text?: string;
            
        }
        /** An interface describing the data returned by calling `border.toJSON()`. */
        export interface BorderData {
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `borderUniversal.toJSON()`. */
        export interface BorderUniversalData {
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `borderCollection.toJSON()`. */
        export interface BorderCollectionData {
            items?: Word.Interfaces.BorderData[];
        }
        /** An interface describing the data returned by calling `borderUniversalCollection.toJSON()`. */
        export interface BorderUniversalCollectionData {
            items?: Word.Interfaces.BorderUniversalData[];
        }
        /** An interface describing the data returned by calling `break.toJSON()`. */
        export interface BreakData {
            
            
        }
        /** An interface describing the data returned by calling `breakCollection.toJSON()`. */
        export interface BreakCollectionData {
            items?: Word.Interfaces.BreakData[];
        }
        /** An interface describing the data returned by calling `buildingBlock.toJSON()`. */
        export interface BuildingBlockData {
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `buildingBlockCategory.toJSON()`. */
        export interface BuildingBlockCategoryData {
            
            
        }
        /** An interface describing the data returned by calling `buildingBlockTypeItem.toJSON()`. */
        export interface BuildingBlockTypeItemData {
            
            
        }
        /** An interface describing the data returned by calling `checkboxContentControl.toJSON()`. */
        export interface CheckboxContentControlData {
            
        }
        /** An interface describing the data returned by calling `coauthoringLock.toJSON()`. */
        export interface CoauthoringLockData {
            
            
            
        }
        /** An interface describing the data returned by calling `coauthoringLockCollection.toJSON()`. */
        export interface CoauthoringLockCollectionData {
            items?: Word.Interfaces.CoauthoringLockData[];
        }
        /** An interface describing the data returned by calling `coauthor.toJSON()`. */
        export interface CoauthorData {
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `coauthorCollection.toJSON()`. */
        export interface CoauthorCollectionData {
            items?: Word.Interfaces.CoauthorData[];
        }
        /** An interface describing the data returned by calling `coauthoring.toJSON()`. */
        export interface CoauthoringData {
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `coauthoringUpdate.toJSON()`. */
        export interface CoauthoringUpdateData {
            
        }
        /** An interface describing the data returned by calling `coauthoringUpdateCollection.toJSON()`. */
        export interface CoauthoringUpdateCollectionData {
            items?: Word.Interfaces.CoauthoringUpdateData[];
        }
        /** An interface describing the data returned by calling `comment.toJSON()`. */
        export interface CommentData {
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `commentCollection.toJSON()`. */
        export interface CommentCollectionData {
            items?: Word.Interfaces.CommentData[];
        }
        /** An interface describing the data returned by calling `commentContentRange.toJSON()`. */
        export interface CommentContentRangeData {
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `commentReply.toJSON()`. */
        export interface CommentReplyData {
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `commentReplyCollection.toJSON()`. */
        export interface CommentReplyCollectionData {
            items?: Word.Interfaces.CommentReplyData[];
        }
        /** An interface describing the data returned by calling `conditionalStyle.toJSON()`. */
        export interface ConditionalStyleData {
            
            
            
            
        }
        /** An interface describing the data returned by calling `xmlMapping.toJSON()`. */
        export interface XmlMappingData {
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `customXmlPrefixMappingCollection.toJSON()`. */
        export interface CustomXmlPrefixMappingCollectionData {
            items?: Word.Interfaces.CustomXmlPrefixMappingData[];
        }
        /** An interface describing the data returned by calling `customXmlPrefixMapping.toJSON()`. */
        export interface CustomXmlPrefixMappingData {
            
            
        }
        /** An interface describing the data returned by calling `customXmlSchema.toJSON()`. */
        export interface CustomXmlSchemaData {
            
            
        }
        /** An interface describing the data returned by calling `customXmlSchemaCollection.toJSON()`. */
        export interface CustomXmlSchemaCollectionData {
            items?: Word.Interfaces.CustomXmlSchemaData[];
        }
        /** An interface describing the data returned by calling `customXmlNodeCollection.toJSON()`. */
        export interface CustomXmlNodeCollectionData {
            items?: Word.Interfaces.CustomXmlNodeData[];
        }
        /** An interface describing the data returned by calling `customXmlNode.toJSON()`. */
        export interface CustomXmlNodeData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `contentControl.toJSON()`. */
        export interface ContentControlData {
            
            
            
            /**
            * Gets the collection of `ContentControl` objects in the content control.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            contentControls?: Word.Interfaces.ContentControlData[];
            
            
            
            /**
            * Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontData;
            
            /**
            * Gets the collection of `InlinePicture` objects in the content control. The collection doesn't include floating images.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            inlinePictures?: Word.Interfaces.InlinePictureData[];
            
            /**
            * Gets the collection of `Paragraph` objects in the content control.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            *
            * Important: For requirement sets 1.1 and 1.2, paragraphs in tables wholly contained within this content control aren't returned. From requirement set 1.3, paragraphs in such tables are also returned.
            */
            paragraphs?: Word.Interfaces.ParagraphData[];
            
            
            
            
            /**
             * Specifies the appearance of the content control. The value can be `boundingBox`, `tags`, or `hidden`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            appearance?: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden";
            /**
             * Specifies a value that indicates whether the user can delete the content control. Mutually exclusive with `removeWhenEdited`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            cannotDelete?: boolean;
            /**
             * Specifies a value that indicates whether the user can edit the contents of the content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            cannotEdit?: boolean;
            /**
             * Specifies the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            color?: string;
            /**
             * Gets an integer that represents the content control identifier.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            id?: number;
            /**
             * Specifies the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            placeholderText?: string;
            /**
             * Specifies a value that indicates whether the content control is removed after it's edited. Mutually exclusive with `cannotDelete`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            removeWhenEdited?: boolean;
            /**
             * Specifies the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the `styleBuiltIn` property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: string;
            
            
            /**
             * Specifies a tag to identify a content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            tag?: string;
            /**
             * Gets the text of the content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            text?: string;
            /**
             * Specifies the title for a content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            title?: string;
            /**
             * Gets the content control type. Only rich text, plain text, check box, dropdown list, combo box, building block gallery, date picker, repeating section, picture, and group content controls are supported currently.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            type?: Word.ContentControlType | "Unknown" | "RichTextInline" | "RichTextParagraphs" | "RichTextTableCell" | "RichTextTableRow" | "RichTextTable" | "PlainTextInline" | "PlainTextParagraph" | "Picture" | "BuildingBlockGallery" | "CheckBox" | "ComboBox" | "DropDownList" | "DatePicker" | "RepeatingSection" | "RichText" | "PlainText" | "Group";
        }
        /** An interface describing the data returned by calling `contentControlCollection.toJSON()`. */
        export interface ContentControlCollectionData {
            items?: Word.Interfaces.ContentControlData[];
        }
        /** An interface describing the data returned by calling `contentControlListItem.toJSON()`. */
        export interface ContentControlListItemData {
            
            
            
        }
        /** An interface describing the data returned by calling `contentControlListItemCollection.toJSON()`. */
        export interface ContentControlListItemCollectionData {
            items?: Word.Interfaces.ContentControlListItemData[];
        }
        /** An interface describing the data returned by calling `customProperty.toJSON()`. */
        export interface CustomPropertyData {
            
            
            
        }
        /** An interface describing the data returned by calling `customPropertyCollection.toJSON()`. */
        export interface CustomPropertyCollectionData {
            items?: Word.Interfaces.CustomPropertyData[];
        }
        /** An interface describing the data returned by calling `customXmlPart.toJSON()`. */
        export interface CustomXmlPartData {
            
            
            
            
            
            
            
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
            * Gets the `Body` object of the main document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            body?: Word.Interfaces.BodyData;
            
            /**
            * Gets the collection of `ContentControl` objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            contentControls?: Word.Interfaces.ContentControlData[];
            
            
            
            
            
            
            
            /**
            * Gets the collection of `Section` objects in the document.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            sections?: Word.Interfaces.SectionData[];
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            /**
             * Indicates whether the changes in the document have been saved. A value of `true` indicates that the document hasn't changed since it was saved.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            saved?: boolean;
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `documentCreated.toJSON()`. */
        export interface DocumentCreatedData {
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `documentProperties.toJSON()`. */
        export interface DocumentPropertiesData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `dropDownListContentControl.toJSON()`. */
        export interface DropDownListContentControlData {
        }
        /** An interface describing the data returned by calling `comboBoxContentControl.toJSON()`. */
        export interface ComboBoxContentControlData {
        }
        /** An interface describing the data returned by calling `field.toJSON()`. */
        export interface FieldData {
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `fieldCollection.toJSON()`. */
        export interface FieldCollectionData {
            items?: Word.Interfaces.FieldData[];
        }
        /** An interface describing the data returned by calling `font.toJSON()`. */
        export interface FontData {
            
            
            
            
            
            
            
            
            
            /**
             * Specifies whether the font is bold. `true` if the font is formatted as bold, otherwise, `false`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            bold?: boolean;
            
            /**
             * Specifies the color for the font. You can provide the value in the '#RRGGBB' format or the color name.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            color?: string;
            
            
            
            
            
            /**
             * Specifies whether the font has a double strikethrough. `true` if the font is formatted as double strikethrough text, otherwise, `false`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            doubleStrikeThrough?: boolean;
            
            
            
            
            /**
             * Specifies the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to `null`. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or `null` for no highlight color. Note: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            highlightColor?: string;
            /**
             * Specifies whether the font is italicized. `true` if the font is italicized, otherwise, `false`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            italic?: boolean;
            
            
            
            /**
             * Specifies the name of the font.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            name?: string;
            
            
            
            
            
            
            
            
            
            
            /**
             * Specifies the font size in points.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            size?: number;
            
            
            
            /**
             * Specifies whether the font has a strikethrough. `true` if the font is formatted as strikethrough text, otherwise, `false`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            strikeThrough?: boolean;
            
            /**
             * Specifies whether the font is a subscript. `true` if the font is formatted as subscript, otherwise, `false`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            subscript?: boolean;
            /**
             * Specifies whether the font is a superscript. `true` if the font is formatted as superscript, otherwise, `false`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            superscript?: boolean;
            /**
             * Specifies the font's underline type. `none` if the font isn't underlined.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            underline?: Word.UnderlineType | "Mixed" | "None" | "Hidden" | "DotLine" | "Single" | "Word" | "Double" | "Thick" | "Dotted" | "DottedHeavy" | "DashLine" | "DashLineHeavy" | "DashLineLong" | "DashLineLongHeavy" | "DotDashLine" | "DotDashLineHeavy" | "TwoDotDashLine" | "TwoDotDashLineHeavy" | "Wave" | "WaveHeavy" | "WaveDouble";
            
        }
        /** An interface describing the data returned by calling `headingStyle.toJSON()`. */
        export interface HeadingStyleData {
            
            
        }
        /** An interface describing the data returned by calling `headingStyleCollection.toJSON()`. */
        export interface HeadingStyleCollectionData {
            items?: Word.Interfaces.HeadingStyleData[];
        }
        /** An interface describing the data returned by calling `hyperlink.toJSON()`. */
        export interface HyperlinkData {
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `hyperlinkCollection.toJSON()`. */
        export interface HyperlinkCollectionData {
            items?: Word.Interfaces.HyperlinkData[];
        }
        /** An interface describing the data returned by calling `inlinePicture.toJSON()`. */
        export interface InlinePictureData {
            /**
             * Specifies a string that represents the alternative text associated with the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            altTextDescription?: string;
            /**
             * Specifies a string that contains the title for the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            altTextTitle?: string;
            /**
             * Specifies a number that describes the height of the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            height?: number;
            /**
             * Specifies a hyperlink on the image. Use a '#' to separate the address part from the optional location part.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            hyperlink?: string;
            
            /**
             * Specifies a value that indicates whether the inline image retains its original proportions when you resize it.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lockAspectRatio?: boolean;
            /**
             * Specifies a number that describes the width of the inline image.
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
        /** An interface describing the data returned by calling `linkFormat.toJSON()`. */
        export interface LinkFormatData {
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `list.toJSON()`. */
        export interface ListData {
            
            
            
            
        }
        /** An interface describing the data returned by calling `listCollection.toJSON()`. */
        export interface ListCollectionData {
            items?: Word.Interfaces.ListData[];
        }
        /** An interface describing the data returned by calling `listItem.toJSON()`. */
        export interface ListItemData {
            
            
            
        }
        /** An interface describing the data returned by calling `listLevel.toJSON()`. */
        export interface ListLevelData {
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `listLevelCollection.toJSON()`. */
        export interface ListLevelCollectionData {
            items?: Word.Interfaces.ListLevelData[];
        }
        /** An interface describing the data returned by calling `listTemplate.toJSON()`. */
        export interface ListTemplateData {
            
            
            
        }
        /** An interface describing the data returned by calling `noteItem.toJSON()`. */
        export interface NoteItemData {
            
            
            
        }
        /** An interface describing the data returned by calling `noteItemCollection.toJSON()`. */
        export interface NoteItemCollectionData {
            items?: Word.Interfaces.NoteItemData[];
        }
        /** An interface describing the data returned by calling `oleFormat.toJSON()`. */
        export interface OleFormatData {
            
        /** An interface describing the data returned by calling `page.toJSON()`. */
        export interface PageData {
            
            
            
            
        }
        /** An interface describing the data returned by calling `pageCollection.toJSON()`. */
        export interface PageCollectionData {
            items?: Word.Interfaces.PageData[];
        }
        /** An interface describing the data returned by calling `pane.toJSON()`. */
        export interface PaneData {
            
            
        }
        /** An interface describing the data returned by calling `paneCollection.toJSON()`. */
        export interface PaneCollectionData {
            items?: Word.Interfaces.PaneData[];
        }
        /** An interface describing the data returned by calling `window.toJSON()`. */
        export interface WindowData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `windowCollection.toJSON()`. */
        export interface WindowCollectionData {
            items?: Word.Interfaces.WindowData[];
        }
        /** An interface describing the data returned by calling `paragraph.toJSON()`. */
        export interface ParagraphData {
            
            
            /**
            * Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontData;
            /**
            * Gets the collection of `InlinePicture` objects in the paragraph. The collection doesn't include floating images.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            inlinePictures?: Word.Interfaces.InlinePictureData[];
            
            
            
            
            /**
             * Specifies the alignment for a paragraph. The value can be `left`, `centered`, `right`, or `justified`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            alignment?: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
            /**
             * Specifies the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            firstLineIndent?: number;
            
            
            /**
             * Specifies the left indent value, in points, for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            leftIndent?: number;
            /**
             * Specifies the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineSpacing?: number;
            /**
             * Specifies the amount of spacing, in grid lines, after the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineUnitAfter?: number;
            /**
             * Specifies the amount of spacing, in grid lines, before the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineUnitBefore?: number;
            /**
             * Specifies the outline level for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            outlineLevel?: number;
            /**
             * Specifies the right indent value, in points, for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            rightIndent?: number;
            /**
             * Specifies the spacing, in points, after the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            spaceAfter?: number;
            /**
             * Specifies the spacing, in points, before the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            spaceBefore?: number;
            /**
             * Specifies the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: string;
            
            
            /**
             * Gets the text of the paragraph.
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
        /** An interface describing the data returned by calling `paragraphFormat.toJSON()`. */
        export interface ParagraphFormatData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `range.toJSON()`. */
        export interface RangeData {
            
            
            
            /**
            * Gets the text format of the range. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontData;
            
            
            /**
            * Gets the collection of `InlinePicture` objects in the range.
            *
            * @remarks
            * [Api set: WordApi 1.2]
            */
            inlinePictures?: Word.Interfaces.InlinePictureData[];
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            /**
             * Specifies the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: string;
            
            /**
             * Gets the text of the range.
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
             * If provided, specifies whether to ignore all punctuation characters between words. The default is `false`. Corresponds to the _Ignore punctuation characters_ check box in the **Find and Replace** dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            ignorePunct?: boolean;
            /**
             * If provided, specifies whether to ignore all whitespace between words. The default is `false`. Corresponds to the _Ignore white-space characters_ check box in the **Find and Replace** dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            ignoreSpace?: boolean;
            /**
             * If provided, specifies whether to perform a case sensitive search. The default is `false`. Corresponds to the _Match case_ check box in the **Find and Replace** dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchCase?: boolean;
            /**
             * If provided, specifies whether to match words that begin with the search string. The default is `false`. Corresponds to the _Match prefix_ check box in the **Find and Replace** dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchPrefix?: boolean;
            /**
             * If provided, specifies whether to match words that end with the search string. The default is `false`. Corresponds to the _Match suffix_ check box in the **Find and Replace** dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchSuffix?: boolean;
            /**
             * If provided, specifies whether to find only entire words, not text that's part of a larger word. The default is `false`. Corresponds to the _Find whole words only_ check box in the **Find and Replace** dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchWholeWord?: boolean;
            /**
             * If provided, specifies whether the search will be performed using special search operators. The default is `false`. Corresponds to the _Use wildcards_ check box in the **Find and Replace** dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchWildcards?: boolean;
        }
        /** An interface describing the data returned by calling `section.toJSON()`. */
        export interface SectionData {
            /**
            * Gets the `Body` object of the section. This doesn't include the header, footer, and other section metadata.
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
            
            
        }
        /** An interface describing the data returned by calling `settingCollection.toJSON()`. */
        export interface SettingCollectionData {
            items?: Word.Interfaces.SettingData[];
        }
        /** An interface describing the data returned by calling `styleCollection.toJSON()`. */
        export interface StyleCollectionData {
            items?: Word.Interfaces.StyleData[];
        }
        /** An interface describing the data returned by calling `style.toJSON()`. */
        export interface StyleData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `shading.toJSON()`. */
        export interface ShadingData {
            
            
            
        }
        /** An interface describing the data returned by calling `shadingUniversal.toJSON()`. */
        export interface ShadingUniversalData {
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `table.toJSON()`. */
        export interface TableData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `tableStyle.toJSON()`. */
        export interface TableStyleData {
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `tabStop.toJSON()`. */
        export interface TabStopData {
            
            
            
            
        }
        /** An interface describing the data returned by calling `tabStopCollection.toJSON()`. */
        export interface TabStopCollectionData {
            items?: Word.Interfaces.TabStopData[];
        }
        /** An interface describing the data returned by calling `tableCollection.toJSON()`. */
        export interface TableCollectionData {
            items?: Word.Interfaces.TableData[];
        }
        /** An interface describing the data returned by calling `tableColumn.toJSON()`. */
        export interface TableColumnData {
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `tableColumnCollection.toJSON()`. */
        export interface TableColumnCollectionData {
            items?: Word.Interfaces.TableColumnData[];
        }
        /** An interface describing the data returned by calling `tableOfAuthorities.toJSON()`. */
        export interface TableOfAuthoritiesData {
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `tableOfAuthoritiesCollection.toJSON()`. */
        export interface TableOfAuthoritiesCollectionData {
            items?: Word.Interfaces.TableOfAuthoritiesData[];
        }
        /** An interface describing the data returned by calling `tableOfAuthoritiesCategory.toJSON()`. */
        export interface TableOfAuthoritiesCategoryData {
            
        }
        /** An interface describing the data returned by calling `tableOfAuthoritiesCategoryCollection.toJSON()`. */
        export interface TableOfAuthoritiesCategoryCollectionData {
            items?: Word.Interfaces.TableOfAuthoritiesCategoryData[];
        }
        /** An interface describing the data returned by calling `tableOfContents.toJSON()`. */
        export interface TableOfContentsData {
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `tableOfContentsCollection.toJSON()`. */
        export interface TableOfContentsCollectionData {
            items?: Word.Interfaces.TableOfContentsData[];
        }
        /** An interface describing the data returned by calling `tableOfFigures.toJSON()`. */
        export interface TableOfFiguresData {
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `tableOfFiguresCollection.toJSON()`. */
        export interface TableOfFiguresCollectionData {
            items?: Word.Interfaces.TableOfFiguresData[];
        }
        /** An interface describing the data returned by calling `tableRow.toJSON()`. */
        export interface TableRowData {
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `tableRowCollection.toJSON()`. */
        export interface TableRowCollectionData {
            items?: Word.Interfaces.TableRowData[];
        }
        /** An interface describing the data returned by calling `tableCell.toJSON()`. */
        export interface TableCellData {
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `tableCellCollection.toJSON()`. */
        export interface TableCellCollectionData {
            items?: Word.Interfaces.TableCellData[];
        }
        /** An interface describing the data returned by calling `tableBorder.toJSON()`. */
        export interface TableBorderData {
            
            
            
        }
        /** An interface describing the data returned by calling `template.toJSON()`. */
        export interface TemplateData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `templateCollection.toJSON()`. */
        export interface TemplateCollectionData {
            items?: Word.Interfaces.TemplateData[];
        }
        /** An interface describing the data returned by calling `trackedChange.toJSON()`. */
        export interface TrackedChangeData {
            
            
            
            
        }
        /** An interface describing the data returned by calling `trackedChangeCollection.toJSON()`. */
        export interface TrackedChangeCollectionData {
            items?: Word.Interfaces.TrackedChangeData[];
        }
        /** An interface describing the data returned by calling `view.toJSON()`. */
        export interface ViewData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `shape.toJSON()`. */
        export interface ShapeData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `shapeGroup.toJSON()`. */
        export interface ShapeGroupData {
            
            
            
        }
        /** An interface describing the data returned by calling `canvas.toJSON()`. */
        export interface CanvasData {
            
            
            
        }
        /** An interface describing the data returned by calling `shapeCollection.toJSON()`. */
        export interface ShapeCollectionData {
            items?: Word.Interfaces.ShapeData[];
        }
        /** An interface describing the data returned by calling `shapeFill.toJSON()`. */
        export interface ShapeFillData {
            
            
            
            
        }
        /** An interface describing the data returned by calling `textFrame.toJSON()`. */
        export interface TextFrameData {
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `shapeTextWrap.toJSON()`. */
        export interface ShapeTextWrapData {
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `reviewer.toJSON()`. */
        export interface ReviewerData {
            
        }
        /** An interface describing the data returned by calling `reviewerCollection.toJSON()`. */
        export interface ReviewerCollectionData {
            items?: Word.Interfaces.ReviewerData[];
        }
        /** An interface describing the data returned by calling `revisionsFilter.toJSON()`. */
        export interface RevisionsFilterData {
            
            
        }
        /** An interface describing the data returned by calling `repeatingSectionItem.toJSON()`. */
        export interface RepeatingSectionItemData {
            
        }
        /** An interface describing the data returned by calling `revision.toJSON()`. */
        export interface RevisionData {
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `revisionCollection.toJSON()`. */
        export interface RevisionCollectionData {
            items?: Word.Interfaces.RevisionData[];
        }
        /** An interface describing the data returned by calling `datePickerContentControl.toJSON()`. */
        export interface DatePickerContentControlData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `pictureContentControl.toJSON()`. */
        export interface PictureContentControlData {
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `groupContentControl.toJSON()`. */
        export interface GroupContentControlData {
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `buildingBlockGalleryContentControl.toJSON()`. */
        export interface BuildingBlockGalleryContentControlData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `repeatingSectionContentControl.toJSON()`. */
        export interface RepeatingSectionContentControlData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `readabilityStatistic.toJSON()`. */
        export interface ReadabilityStatisticData {
            
            
        }
        /** An interface describing the data returned by calling `readabilityStatisticCollection.toJSON()`. */
        export interface ReadabilityStatisticCollectionData {
            items?: Word.Interfaces.ReadabilityStatisticData[];
        }
        /** An interface describing the data returned by calling `webSettings.toJSON()`. */
        export interface WebSettingsData {
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `xmlNodeCollection.toJSON()`. */
        export interface XmlNodeCollectionData {
            items?: Word.Interfaces.XmlNodeData[];
        }
        /** An interface describing the data returned by calling `xmlNode.toJSON()`. */
        export interface XmlNodeData {
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `htmlDivision.toJSON()`. */
        export interface HtmlDivisionData {
            
            
            
            
        }
        /** An interface describing the data returned by calling `htmlDivisionCollection.toJSON()`. */
        export interface HtmlDivisionCollectionData {
            items?: Word.Interfaces.HtmlDivisionData[];
        }
        /** An interface describing the data returned by calling `frame.toJSON()`. */
        export interface FrameData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `frameCollection.toJSON()`. */
        export interface FrameCollectionData {
            items?: Word.Interfaces.FrameData[];
        }
        /** An interface describing the data returned by calling `documentLibraryVersion.toJSON()`. */
        export interface DocumentLibraryVersionData {
            
            
            
        }
        /** An interface describing the data returned by calling `documentLibraryVersionCollection.toJSON()`. */
        export interface DocumentLibraryVersionCollectionData {
            items?: Word.Interfaces.DocumentLibraryVersionData[];
        }
        /** An interface describing the data returned by calling `dropCap.toJSON()`. */
        export interface DropCapData {
            
            
            
            
        }
        /** An interface describing the data returned by calling `listFormat.toJSON()`. */
        export interface ListFormatData {
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `fillFormat.toJSON()`. */
        export interface FillFormatData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `glowFormat.toJSON()`. */
        export interface GlowFormatData {
            
            
            
        }
        /** An interface describing the data returned by calling `lineFormat.toJSON()`. */
        export interface LineFormatData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `reflectionFormat.toJSON()`. */
        export interface ReflectionFormatData {
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `colorFormat.toJSON()`. */
        export interface ColorFormatData {
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `shadowFormat.toJSON()`. */
        export interface ShadowFormatData {
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `threeDimensionalFormat.toJSON()`. */
        export interface ThreeDimensionalFormatData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `bibliography.toJSON()`. */
        export interface BibliographyData {
            
            
        }
        /** An interface describing the data returned by calling `sourceCollection.toJSON()`. */
        export interface SourceCollectionData {
            items?: Word.Interfaces.SourceData[];
        }
        /** An interface describing the data returned by calling `source.toJSON()`. */
        export interface SourceData {
            
            
            
        }
        /** An interface describing the data returned by calling `pageSetup.toJSON()`. */
        export interface PageSetupData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `lineNumbering.toJSON()`. */
        export interface LineNumberingData {
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `textColumnCollection.toJSON()`. */
        export interface TextColumnCollectionData {
            items?: Word.Interfaces.TextColumnData[];
        }
        /** An interface describing the data returned by calling `textColumn.toJSON()`. */
        export interface TextColumnData {
            
            
        }
        /** An interface describing the data returned by calling `selection.toJSON()`. */
        export interface SelectionData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `rangeScopedCollection.toJSON()`. */
        export interface RangeScopedCollectionData {
            items?: Word.Interfaces.RangeData[];
        }
        /** An interface describing the data returned by calling `bookmark.toJSON()`. */
        export interface BookmarkData {
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `bookmarkCollection.toJSON()`. */
        export interface BookmarkCollectionData {
            items?: Word.Interfaces.BookmarkData[];
        }
        /** An interface describing the data returned by calling `index.toJSON()`. */
        export interface IndexData {
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `indexCollection.toJSON()`. */
        export interface IndexCollectionData {
            items?: Word.Interfaces.IndexData[];
        }
        /** An interface describing the data returned by calling `listTemplateCollection.toJSON()`. */
        export interface ListTemplateCollectionData {
            items?: Word.Interfaces.ListTemplateData[];
        }
        /** An interface describing the data returned by calling `listTemplateGallery.toJSON()`. */
        export interface ListTemplateGalleryData {
        }
        /** An interface describing the data returned by calling `listTemplateGalleryCollection.toJSON()`. */
        export interface ListTemplateGalleryCollectionData {
            items?: Word.Interfaces.ListTemplateGalleryData[];
        }
        
        
        
        
        
        
        
        /**
         * Represents the body of a document or a section.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        export interface BodyLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Gets the text format of the body. Use this to get and set font name, size, color, and other properties.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            font?: Word.Interfaces.FontLoadOptions;
            
            
            /**
            * Gets the content control that contains the body. Throws an `ItemNotFound` error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            
            
            
            /**
             * Specifies the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: boolean;
            
            /**
             * Gets the text of the body. Use the insertText method to insert text.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            text?: boolean;
            
        }
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        /**
         * Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text, plain text, checkbox, dropdown list, and combo box content controls are supported.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        export interface ContentControlLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
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
            * Gets the content control that contains the content control. Throws an `ItemNotFound` error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            
            
            
            
            
            
            
            
            /**
             * Specifies the appearance of the content control. The value can be `boundingBox`, `tags`, or `hidden`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            appearance?: boolean;
            /**
             * Specifies a value that indicates whether the user can delete the content control. Mutually exclusive with `removeWhenEdited`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            cannotDelete?: boolean;
            /**
             * Specifies a value that indicates whether the user can edit the contents of the content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            cannotEdit?: boolean;
            /**
             * Specifies the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            color?: boolean;
            /**
             * Gets an integer that represents the content control identifier.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            id?: boolean;
            /**
             * Specifies the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            placeholderText?: boolean;
            /**
             * Specifies a value that indicates whether the content control is removed after it's edited. Mutually exclusive with `cannotDelete`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            removeWhenEdited?: boolean;
            /**
             * Specifies the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the `styleBuiltIn` property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: boolean;
            
            
            /**
             * Specifies a tag to identify a content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            tag?: boolean;
            /**
             * Gets the text of the content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            text?: boolean;
            /**
             * Specifies the title for a content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            title?: boolean;
            /**
             * Gets the content control type. Only rich text, plain text, check box, dropdown list, combo box, building block gallery, date picker, repeating section, picture, and group content controls are supported currently.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            type?: boolean;
        }
        /**
         * Contains a collection of {@link Word.ContentControl} objects. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text, plain text, checkbox, dropdown list, and combo box content controls are supported.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        export interface ContentControlCollectionLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
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
            * For EACH ITEM in the collection: Gets the content control that contains the content control. Throws an `ItemNotFound` error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            
            
            
            
            
            
            
            
            /**
             * For EACH ITEM in the collection: Specifies the appearance of the content control. The value can be `boundingBox`, `tags`, or `hidden`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            appearance?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies a value that indicates whether the user can delete the content control. Mutually exclusive with `removeWhenEdited`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            cannotDelete?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies a value that indicates whether the user can edit the contents of the content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            cannotEdit?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            color?: boolean;
            /**
             * For EACH ITEM in the collection: Gets an integer that represents the content control identifier.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            id?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            placeholderText?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies a value that indicates whether the content control is removed after it's edited. Mutually exclusive with `cannotDelete`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            removeWhenEdited?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the `styleBuiltIn` property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: boolean;
            
            
            /**
             * For EACH ITEM in the collection: Specifies a tag to identify a content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            tag?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the text of the content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            text?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies the title for a content control.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            title?: boolean;
            /**
             * For EACH ITEM in the collection: Gets the content control type. Only rich text, plain text, check box, dropdown list, combo box, building block gallery, date picker, repeating section, picture, and group content controls are supported currently.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            type?: boolean;
        }
        
        
        
        
        
        
        
        /**
         * The `Document` object is the top level object. A `Document` object contains one or more sections, content controls, and the body that contains the contents of the document.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        export interface DocumentLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            
            
            
            /**
            * Gets the `Body` object of the main document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            body?: Word.Interfaces.BodyLoadOptions;
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            /**
             * Indicates whether the changes in the document have been saved. A value of `true` indicates that the document hasn't changed since it was saved.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            saved?: boolean;
            
            
            
            
            
            
            
            
            
            
            
        }
        
        
        
        
        /**
         * Represents a font.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        export interface FontLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            
            
            
            
            
            
            
            /**
             * Specifies whether the font is bold. `true` if the font is formatted as bold, otherwise, `false`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            bold?: boolean;
            
            /**
             * Specifies the color for the font. You can provide the value in the '#RRGGBB' format or the color name.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            color?: boolean;
            
            
            
            
            
            /**
             * Specifies whether the font has a double strikethrough. `true` if the font is formatted as double strikethrough text, otherwise, `false`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            doubleStrikeThrough?: boolean;
            
            
            
            
            /**
             * Specifies the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to `null`. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or `null` for no highlight color. Note: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            highlightColor?: boolean;
            /**
             * Specifies whether the font is italicized. `true` if the font is italicized, otherwise, `false`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            italic?: boolean;
            
            
            
            /**
             * Specifies the name of the font.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            name?: boolean;
            
            
            
            
            
            
            
            
            
            
            /**
             * Specifies the font size in points.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            size?: boolean;
            
            
            
            /**
             * Specifies whether the font has a strikethrough. `true` if the font is formatted as strikethrough text, otherwise, `false`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            strikeThrough?: boolean;
            
            /**
             * Specifies whether the font is a subscript. `true` if the font is formatted as subscript, otherwise, `false`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            subscript?: boolean;
            /**
             * Specifies whether the font is a superscript. `true` if the font is formatted as superscript, otherwise, `false`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            superscript?: boolean;
            /**
             * Specifies the font's underline type. `none` if the font isn't underlined.
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
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
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
            * Gets the content control that contains the inline image. Throws an `ItemNotFound` error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            
            
            
            
            
            /**
             * Specifies a string that represents the alternative text associated with the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            altTextDescription?: boolean;
            /**
             * Specifies a string that contains the title for the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            altTextTitle?: boolean;
            /**
             * Specifies a number that describes the height of the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            height?: boolean;
            /**
             * Specifies a hyperlink on the image. Use a '#' to separate the address part from the optional location part.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            hyperlink?: boolean;
            
            /**
             * Specifies a value that indicates whether the inline image retains its original proportions when you resize it.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lockAspectRatio?: boolean;
            /**
             * Specifies a number that describes the width of the inline image.
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
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
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
            * For EACH ITEM in the collection: Gets the content control that contains the inline image. Throws an `ItemNotFound` error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            
            
            
            
            
            /**
             * For EACH ITEM in the collection: Specifies a string that represents the alternative text associated with the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            altTextDescription?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies a string that contains the title for the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            altTextTitle?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies a number that describes the height of the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            height?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies a hyperlink on the image. Use a '#' to separate the address part from the optional location part.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            hyperlink?: boolean;
            
            /**
             * For EACH ITEM in the collection: Specifies a value that indicates whether the inline image retains its original proportions when you resize it.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lockAspectRatio?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies a number that describes the width of the inline image.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            width?: boolean;
        }
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        /**
         * Represents a single paragraph in a selection, range, content control, or document body.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        export interface ParagraphLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
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
            * Gets the content control that contains the paragraph. Throws an `ItemNotFound` error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            
            
            
            
            
            
            /**
             * Specifies the alignment for a paragraph. The value can be `left`, `centered`, `right`, or `justified`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            alignment?: boolean;
            /**
             * Specifies the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            firstLineIndent?: boolean;
            
            
            /**
             * Specifies the left indent value, in points, for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            leftIndent?: boolean;
            /**
             * Specifies the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineSpacing?: boolean;
            /**
             * Specifies the amount of spacing, in grid lines, after the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineUnitAfter?: boolean;
            /**
             * Specifies the amount of spacing, in grid lines, before the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineUnitBefore?: boolean;
            /**
             * Specifies the outline level for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            outlineLevel?: boolean;
            /**
             * Specifies the right indent value, in points, for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            rightIndent?: boolean;
            /**
             * Specifies the spacing, in points, after the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            spaceAfter?: boolean;
            /**
             * Specifies the spacing, in points, before the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            spaceBefore?: boolean;
            /**
             * Specifies the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: boolean;
            
            
            /**
             * Gets the text of the paragraph.
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
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
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
            * For EACH ITEM in the collection: Gets the content control that contains the paragraph. Throws an `ItemNotFound` error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            
            
            
            
            
            
            /**
             * For EACH ITEM in the collection: Specifies the alignment for a paragraph. The value can be `left`, `centered`, `right`, or `justified`.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            alignment?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            firstLineIndent?: boolean;
            
            
            /**
             * For EACH ITEM in the collection: Specifies the left indent value, in points, for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            leftIndent?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineSpacing?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies the amount of spacing, in grid lines, after the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineUnitAfter?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies the amount of spacing, in grid lines, before the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            lineUnitBefore?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies the outline level for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            outlineLevel?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies the right indent value, in points, for the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            rightIndent?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies the spacing, in points, after the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            spaceAfter?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies the spacing, in points, before the paragraph.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            spaceBefore?: boolean;
            /**
             * For EACH ITEM in the collection: Specifies the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: boolean;
            
            
            /**
             * For EACH ITEM in the collection: Gets the text of the paragraph.
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
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
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
            * Gets the currently supported content control that contains the range. Throws an `ItemNotFound` error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            /**
             * Specifies the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: boolean;
            
            /**
             * Gets the text of the range.
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
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
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
            * For EACH ITEM in the collection: Gets the currently supported content control that contains the range. Throws an `ItemNotFound` error if there isn't a parent content control.
            *
            * @remarks
            * [Api set: WordApi 1.1]
            */
            parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            /**
             * For EACH ITEM in the collection: Specifies the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            style?: boolean;
            
            /**
             * For EACH ITEM in the collection: Gets the text of the range.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            text?: boolean;
            
        }
        /**
         * Specifies the options to be included in a search operation.
                    To learn more about how to use search options in the Word JavaScript APIs, read {@link https://learn.microsoft.com/office/dev/add-ins/word/search-option-guidance | Use search options to find text in your Word add-in}.
         *
         * @remarks
         * [Api set: WordApi 1.1]
         */
        export interface SearchOptionsLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * If provided, specifies whether to ignore all punctuation characters between words. The default is `false`. Corresponds to the _Ignore punctuation characters_ check box in the **Find and Replace** dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            ignorePunct?: boolean;
            /**
             * If provided, specifies whether to ignore all whitespace between words. The default is `false`. Corresponds to the _Ignore white-space characters_ check box in the **Find and Replace** dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            ignoreSpace?: boolean;
            /**
             * If provided, specifies whether to perform a case sensitive search. The default is `false`. Corresponds to the _Match case_ check box in the **Find and Replace** dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchCase?: boolean;
            /**
             * If provided, specifies whether to match words that begin with the search string. The default is `false`. Corresponds to the _Match prefix_ check box in the **Find and Replace** dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchPrefix?: boolean;
            /**
             * If provided, specifies whether to match words that end with the search string. The default is `false`. Corresponds to the _Match suffix_ check box in the **Find and Replace** dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchSuffix?: boolean;
            /**
             * If provided, specifies whether to find only entire words, not text that's part of a larger word. The default is `false`. Corresponds to the _Find whole words only_ check box in the **Find and Replace** dialog box.
             *
             * @remarks
             * [Api set: WordApi 1.1]
             */
            matchWholeWord?: boolean;
            /**
             * If provided, specifies whether the search will be performed using special search operators. The default is `false`. Corresponds to the _Use wildcards_ check box in the **Find and Replace** dialog box.
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
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Gets the `Body` object of the section. This doesn't include the header, footer, and other section metadata.
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
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * For EACH ITEM in the collection: Gets the `Body` object of the section. This doesn't include the header, footer, and other section metadata.
            *
            * @remarks
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
     * @param objects - An array of previously created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by `context.sync()`.
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of `context.sync()`). The context parameter facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the RequestContext is required to get access to the Word object model from the add-in.
     */
    export function run<T>(objects: OfficeExtension.ClientObject[], batch: (context: Word.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the Word object model, using the RequestContext of a previously created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param object - A previously created API object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by `context.sync()`.
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of `context.sync()`). The context parameter facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the RequestContext is required to get access to the Word object model from the add-in.
     */
    export function run<T>(object: OfficeExtension.ClientObject, batch: (context: Word.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the Word object model, using a new RequestContext. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of `context.sync()`). The context parameter facilitates requests to the Word application. Since the Office add-in and the Word application run in two different processes, the RequestContext is required to get access to the Word object model from the add-in.
     */
    export function run<T>(batch: (context: Word.RequestContext) => Promise<T>): Promise<T>;
}


////////////////////////////////////////////////////////////////
//////////////////////// End Word APIs /////////////////////////
////////////////////////////////////////////////////////////////