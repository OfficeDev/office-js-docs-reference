import { OfficeExtension } from "../api-extractor-inputs-office/office"
////////////////////////////////////////////////////////////////
////////////////////// Begin OneNote APIs //////////////////////
////////////////////////////////////////////////////////////////

export declare namespace OneNote {
    /**
     *
     * Represents the top-level object that contains all globally addressable OneNote objects such as notebooks, the active notebook, and the active section.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class Application extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the collection of notebooks that are open in the OneNote application instance. In OneNote Online, only one notebook at a time is open in the application instance. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly notebooks: OneNote.NotebookCollection;
        /**
         *
         * Gets the active notebook if one exists. If no notebook is active, throws ItemNotFound.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActiveNotebook(): OneNote.Notebook;
        /**
         *
         * Gets the active notebook if one exists. If no notebook is active, returns null.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActiveNotebookOrNull(): OneNote.Notebook;
        /**
         *
         * Gets the active outline if one exists, If no outline is active, throws ItemNotFound.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActiveOutline(): OneNote.Outline;
        /**
         *
         * Gets the active outline if one exists, otherwise returns null.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActiveOutlineOrNull(): OneNote.Outline;
        /**
         *
         * Gets the active page if one exists. If no page is active, throws ItemNotFound.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActivePage(): OneNote.Page;
        /**
         *
         * Gets the active page if one exists. If no page is active, returns null.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActivePageOrNull(): OneNote.Page;
        /**
         *
         * Gets the active Paragraph if one exists, If no Paragraph is active, throws ItemNotFound.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActiveParagraph(): OneNote.Paragraph;
        /**
         *
         * Gets the active Paragraph if one exists, otherwise returns null.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActiveParagraphOrNull(): OneNote.Paragraph;
        /**
         *
         * Gets the active section if one exists. If no section is active, throws ItemNotFound.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActiveSection(): OneNote.Section;
        /**
         *
         * Gets the active section if one exists. If no section is active, returns null.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getActiveSectionOrNull(): OneNote.Section;
        getWindowSize(): OfficeExtension.ClientResult<number[]>;
        insertHtmlAtCurrentPosition(html: string): void;
        isViewingDeletedNotes(): OfficeExtension.ClientResult<boolean>;
        /**
         *
         * Opens the specified page in the application instance.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param page - The page to open.
         */
        navigateToPage(page: OneNote.Page): void;
        /**
         *
         * Gets the specified page, and opens it in the application instance.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param url - The client url of the page to open.
         */
        navigateToPageWithClientUrl(url: string): OneNote.Page;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.Application` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.Application` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.Application` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.ApplicationLoadOptions): OneNote.Application;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.Application;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): OneNote.Application;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OneNote.Application object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.ApplicationData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): OneNote.Interfaces.ApplicationData;
    }
    /**
     *
     * Represents ink analysis data for a given set of ink strokes.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkAnalysis extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the parent page object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly page: OneNote.Page;
        /**
         *
         * Gets the ID of the InkAnalysis object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: OneNote.InkAnalysis): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.InkAnalysisUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: OneNote.InkAnalysis): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.InkAnalysis` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.InkAnalysis` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.InkAnalysis` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.InkAnalysisLoadOptions): OneNote.InkAnalysis;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.InkAnalysis;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): OneNote.InkAnalysis;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkAnalysis;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkAnalysis;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OneNote.InkAnalysis object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkAnalysisData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): OneNote.Interfaces.InkAnalysisData;
    }
    /**
     *
     * Represents ink analysis data for an identified paragraph formed by ink strokes.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkAnalysisParagraph extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Reference to the parent InkAnalysisPage. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly inkAnalysis: OneNote.InkAnalysis;
        /**
         *
         * Gets the ink analysis lines in this ink analysis paragraph. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly lines: OneNote.InkAnalysisLineCollection;
        /**
         *
         * Gets the ID of the InkAnalysisParagraph object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: OneNote.InkAnalysisParagraph): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.InkAnalysisParagraphUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: OneNote.InkAnalysisParagraph): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.InkAnalysisParagraph` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.InkAnalysisParagraph` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.InkAnalysisParagraph` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.InkAnalysisParagraphLoadOptions): OneNote.InkAnalysisParagraph;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.InkAnalysisParagraph;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): OneNote.InkAnalysisParagraph;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkAnalysisParagraph;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkAnalysisParagraph;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OneNote.InkAnalysisParagraph object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkAnalysisParagraphData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): OneNote.Interfaces.InkAnalysisParagraphData;
    }
    /**
     *
     * Represents a collection of InkAnalysisParagraph objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkAnalysisParagraphCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: OneNote.InkAnalysisParagraph[];
        /**
         *
         * Returns the number of InkAnalysisParagraphs in the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        /**
         *
         * Gets a InkAnalysisParagraph object by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - The ID of the InkAnalysisParagraph object, or the index location of the InkAnalysisParagraph object in the collection.
         */
        getItem(index: number | string): OneNote.InkAnalysisParagraph;
        /**
         *
         * Gets a InkAnalysisParagraph on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.InkAnalysisParagraph;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.InkAnalysisParagraphCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.InkAnalysisParagraphCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.InkAnalysisParagraphCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.InkAnalysisParagraphCollectionLoadOptions & OneNote.Interfaces.CollectionLoadOptions): OneNote.InkAnalysisParagraphCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.InkAnalysisParagraphCollection;
        load(option?: OfficeExtension.LoadOption): OneNote.InkAnalysisParagraphCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkAnalysisParagraphCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkAnalysisParagraphCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `OneNote.InkAnalysisParagraphCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkAnalysisParagraphCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): OneNote.Interfaces.InkAnalysisParagraphCollectionData;
    }
    /**
     *
     * Represents ink analysis data for an identified text line formed by ink strokes.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkAnalysisLine extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Reference to the parent InkAnalysisParagraph. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly paragraph: OneNote.InkAnalysisParagraph;
        /**
         *
         * Gets the ink analysis words in this ink analysis line. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly words: OneNote.InkAnalysisWordCollection;
        /**
         *
         * Gets the ID of the InkAnalysisLine object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: OneNote.InkAnalysisLine): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.InkAnalysisLineUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: OneNote.InkAnalysisLine): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.InkAnalysisLine` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.InkAnalysisLine` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.InkAnalysisLine` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.InkAnalysisLineLoadOptions): OneNote.InkAnalysisLine;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.InkAnalysisLine;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): OneNote.InkAnalysisLine;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkAnalysisLine;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkAnalysisLine;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OneNote.InkAnalysisLine object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkAnalysisLineData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): OneNote.Interfaces.InkAnalysisLineData;
    }
    /**
     *
     * Represents a collection of InkAnalysisLine objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkAnalysisLineCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: OneNote.InkAnalysisLine[];
        /**
         *
         * Returns the number of InkAnalysisLines in the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        /**
         *
         * Gets a InkAnalysisLine object by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - The ID of the InkAnalysisLine object, or the index location of the InkAnalysisLine object in the collection.
         */
        getItem(index: number | string): OneNote.InkAnalysisLine;
        /**
         *
         * Gets a InkAnalysisLine on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.InkAnalysisLine;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.InkAnalysisLineCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.InkAnalysisLineCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.InkAnalysisLineCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.InkAnalysisLineCollectionLoadOptions & OneNote.Interfaces.CollectionLoadOptions): OneNote.InkAnalysisLineCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.InkAnalysisLineCollection;
        load(option?: OfficeExtension.LoadOption): OneNote.InkAnalysisLineCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkAnalysisLineCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkAnalysisLineCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `OneNote.InkAnalysisLineCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkAnalysisLineCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): OneNote.Interfaces.InkAnalysisLineCollectionData;
    }
    /**
     *
     * Represents ink analysis data for an identified word formed by ink strokes.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkAnalysisWord extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Reference to the parent InkAnalysisLine. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly line: OneNote.InkAnalysisLine;
        /**
         *
         * Gets the ID of the InkAnalysisWord object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * The id of the recognized language in this inkAnalysisWord. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly languageId: string;
        /**
         *
         * Weak references to the ink strokes that were recognized as part of this ink analysis word. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly strokePointers: OneNote.InkStrokePointer[];
        /**
         *
         * The words that were recognized in this ink word, in order of likelihood. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly wordAlternates: string[];
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: OneNote.InkAnalysisWord): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.InkAnalysisWordUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: OneNote.InkAnalysisWord): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.InkAnalysisWord` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.InkAnalysisWord` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.InkAnalysisWord` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.InkAnalysisWordLoadOptions): OneNote.InkAnalysisWord;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.InkAnalysisWord;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): OneNote.InkAnalysisWord;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkAnalysisWord;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkAnalysisWord;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OneNote.InkAnalysisWord object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkAnalysisWordData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): OneNote.Interfaces.InkAnalysisWordData;
    }
    /**
     *
     * Represents a collection of InkAnalysisWord objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkAnalysisWordCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: OneNote.InkAnalysisWord[];
        /**
         *
         * Returns the number of InkAnalysisWords in the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        /**
         *
         * Gets a InkAnalysisWord object by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - The ID of the InkAnalysisWord object, or the index location of the InkAnalysisWord object in the collection.
         */
        getItem(index: number | string): OneNote.InkAnalysisWord;
        /**
         *
         * Gets a InkAnalysisWord on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.InkAnalysisWord;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.InkAnalysisWordCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.InkAnalysisWordCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.InkAnalysisWordCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.InkAnalysisWordCollectionLoadOptions & OneNote.Interfaces.CollectionLoadOptions): OneNote.InkAnalysisWordCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.InkAnalysisWordCollection;
        load(option?: OfficeExtension.LoadOption): OneNote.InkAnalysisWordCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkAnalysisWordCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkAnalysisWordCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `OneNote.InkAnalysisWordCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkAnalysisWordCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): OneNote.Interfaces.InkAnalysisWordCollectionData;
    }
    /**
     *
     * Represents a group of ink strokes.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class FloatingInk extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the strokes of the FloatingInk object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly inkStrokes: OneNote.InkStrokeCollection;
        /**
         *
         * Gets the PageContent parent of the FloatingInk object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly pageContent: OneNote.PageContent;
        /**
         *
         * Gets the ID of the FloatingInk object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.FloatingInk` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.FloatingInk` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.FloatingInk` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.FloatingInkLoadOptions): OneNote.FloatingInk;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.FloatingInk;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): OneNote.FloatingInk;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.FloatingInk;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.FloatingInk;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OneNote.FloatingInk object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.FloatingInkData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): OneNote.Interfaces.FloatingInkData;
    }
    /**
     *
     * Represents a single stroke of ink.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkStroke extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the ID of the InkStroke object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly floatingInk: OneNote.FloatingInk;
        /**
         *
         * Gets the ID of the InkStroke object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.InkStroke` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.InkStroke` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.InkStroke` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.InkStrokeLoadOptions): OneNote.InkStroke;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.InkStroke;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): OneNote.InkStroke;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkStroke;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkStroke;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OneNote.InkStroke object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkStrokeData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): OneNote.Interfaces.InkStrokeData;
    }
    /**
     *
     * Represents a collection of InkStroke objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkStrokeCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: OneNote.InkStroke[];
        /**
         *
         * Returns the number of InkStrokes in the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        /**
         *
         * Gets a InkStroke object by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - The ID of the InkStroke object, or the index location of the InkStroke object in the collection.
         */
        getItem(index: number | string): OneNote.InkStroke;
        /**
         *
         * Gets a InkStroke on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.InkStroke;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.InkStrokeCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.InkStrokeCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.InkStrokeCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.InkStrokeCollectionLoadOptions & OneNote.Interfaces.CollectionLoadOptions): OneNote.InkStrokeCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.InkStrokeCollection;
        load(option?: OfficeExtension.LoadOption): OneNote.InkStrokeCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkStrokeCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkStrokeCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `OneNote.InkStrokeCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkStrokeCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): OneNote.Interfaces.InkStrokeCollectionData;
    }
    /**
     *
     * A container for the ink in a word in a paragraph.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkWord extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * The parent paragraph containing the ink word. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly paragraph: OneNote.Paragraph;
        /**
         *
         * Gets the ID of the InkWord object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * The id of the recognized language in this ink word. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly languageId: string;
        /**
         *
         * The words that were recognized in this ink word, in order of likelihood. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly wordAlternates: string[];
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.InkWord` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.InkWord` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.InkWord` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.InkWordLoadOptions): OneNote.InkWord;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.InkWord;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): OneNote.InkWord;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkWord;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkWord;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OneNote.InkWord object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkWordData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): OneNote.Interfaces.InkWordData;
    }
    /**
     *
     * Represents a collection of InkWord objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class InkWordCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: OneNote.InkWord[];
        /**
         *
         * Returns the number of InkWords in the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        /**
         *
         * Gets a InkWord object by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - The ID of the InkWord object, or the index location of the InkWord object in the collection.
         */
        getItem(index: number | string): OneNote.InkWord;
        /**
         *
         * Gets a InkWord on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.InkWord;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.InkWordCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.InkWordCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.InkWordCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.InkWordCollectionLoadOptions & OneNote.Interfaces.CollectionLoadOptions): OneNote.InkWordCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.InkWordCollection;
        load(option?: OfficeExtension.LoadOption): OneNote.InkWordCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.InkWordCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.InkWordCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `OneNote.InkWordCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.InkWordCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): OneNote.Interfaces.InkWordCollectionData;
    }
    /**
     *
     * Represents a OneNote notebook. Notebooks contain section groups and sections.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class Notebook extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * The section groups in the notebook. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly sectionGroups: OneNote.SectionGroupCollection;
        /**
         *
         * The the sections of the notebook. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly sections: OneNote.SectionCollection;
        /**
         *
         * The url of the site that this notebook is located. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly baseUrl: string;
        /**
         *
         * The client url of the notebook. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly clientUrl: string;
        /**
         *
         * Gets the ID of the notebook. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * True if the Notebook is not created by the user (i.e. 'Misplaced Sections'). Read only
         *
         * [Api set: OneNoteApi 1.2]
         */
        readonly isVirtual: boolean;
        /**
         *
         * Gets the name of the notebook. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly name: string;
        /**
         *
         * Adds a new section to the end of the notebook.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param name - The name of the new section.
         */
        addSection(name: string): OneNote.Section;
        /**
         *
         * Adds a new section group to the end of the notebook.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param name - The name of the new section.
         */
        addSectionGroup(name: string): OneNote.SectionGroup;
        /**
         *
         * Gets the REST API ID.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getRestApiId(): OfficeExtension.ClientResult<string>;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.Notebook` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.Notebook` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.Notebook` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.NotebookLoadOptions): OneNote.Notebook;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.Notebook;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): OneNote.Notebook;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.Notebook;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.Notebook;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OneNote.Notebook object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.NotebookData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): OneNote.Interfaces.NotebookData;
    }
    /**
     *
     * Represents a collection of notebooks.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class NotebookCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: OneNote.Notebook[];
        /**
         *
         * Returns the number of notebooks in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        /**
         *
         * Gets the collection of notebooks with the specified name that are open in the application instance.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param name - The name of the notebook.
         */
        getByName(name: string): OneNote.NotebookCollection;
        /**
         *
         * Gets a notebook by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - The ID of the notebook, or the index location of the notebook in the collection.
         */
        getItem(index: number | string): OneNote.Notebook;
        /**
         *
         * Gets a notebook on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.Notebook;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.NotebookCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.NotebookCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.NotebookCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.NotebookCollectionLoadOptions & OneNote.Interfaces.CollectionLoadOptions): OneNote.NotebookCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.NotebookCollection;
        load(option?: OfficeExtension.LoadOption): OneNote.NotebookCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.NotebookCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.NotebookCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `OneNote.NotebookCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.NotebookCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): OneNote.Interfaces.NotebookCollectionData;
    }
    /**
     *
     * Represents a OneNote section group. Section groups can contain sections and other section groups.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class SectionGroup extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the notebook that contains the section group. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly notebook: OneNote.Notebook;
        /**
         *
         * Gets the section group that contains the section group. Throws ItemNotFound if the section group is a direct child of the notebook. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentSectionGroup: OneNote.SectionGroup;
        /**
         *
         * Gets the section group that contains the section group. Returns null if the section group is a direct child of the notebook. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentSectionGroupOrNull: OneNote.SectionGroup;
        /**
         *
         * The collection of section groups in the section group. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly sectionGroups: OneNote.SectionGroupCollection;
        /**
         *
         * The collection of sections in the section group. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly sections: OneNote.SectionCollection;
        /**
         *
         * The client url of the section group. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly clientUrl: string;
        /**
         *
         * Gets the ID of the section group. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * Gets the name of the section group. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly name: string;
        /**
         *
         * Adds a new section to the end of the section group.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param title - The name of the new section.
         */
        addSection(title: string): OneNote.Section;
        /**
         *
         * Adds a new section group to the end of this sectionGroup.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param name - The name of the new section.
         */
        addSectionGroup(name: string): OneNote.SectionGroup;
        /**
         *
         * Gets the REST API ID.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getRestApiId(): OfficeExtension.ClientResult<string>;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.SectionGroup` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.SectionGroup` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.SectionGroup` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.SectionGroupLoadOptions): OneNote.SectionGroup;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.SectionGroup;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): OneNote.SectionGroup;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.SectionGroup;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.SectionGroup;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OneNote.SectionGroup object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.SectionGroupData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): OneNote.Interfaces.SectionGroupData;
    }
    /**
     *
     * Represents a collection of section groups.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class SectionGroupCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: OneNote.SectionGroup[];
        /**
         *
         * Returns the number of section groups in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        /**
         *
         * Gets the collection of section groups with the specified name.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param name - The name of the section group.
         */
        getByName(name: string): OneNote.SectionGroupCollection;
        /**
         *
         * Gets a section group by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - The ID of the section group, or the index location of the section group in the collection.
         */
        getItem(index: number | string): OneNote.SectionGroup;
        /**
         *
         * Gets a section group on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.SectionGroup;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.SectionGroupCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.SectionGroupCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.SectionGroupCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.SectionGroupCollectionLoadOptions & OneNote.Interfaces.CollectionLoadOptions): OneNote.SectionGroupCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.SectionGroupCollection;
        load(option?: OfficeExtension.LoadOption): OneNote.SectionGroupCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.SectionGroupCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.SectionGroupCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `OneNote.SectionGroupCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.SectionGroupCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): OneNote.Interfaces.SectionGroupCollectionData;
    }
    /**
     *
     * Represents a OneNote section. Sections can contain pages.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class Section extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the notebook that contains the section. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly notebook: OneNote.Notebook;
        /**
         *
         * The collection of pages in the section. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly pages: OneNote.PageCollection;
        /**
         *
         * Gets the section group that contains the section. Throws ItemNotFound if the section is a direct child of the notebook. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentSectionGroup: OneNote.SectionGroup;
        /**
         *
         * Gets the section group that contains the section. Returns null if the section is a direct child of the notebook. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentSectionGroupOrNull: OneNote.SectionGroup;
        /**
         *
         * The client url of the section. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly clientUrl: string;
        /**
         *
         * Gets the ID of the section. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * True if this section is encrypted with a password. Read only
         *
         * [Api set: OneNoteApi 1.2]
         */
        readonly isEncrypted: boolean;
        /**
         *
         * True if this section is locked. Read only
         *
         * [Api set: OneNoteApi 1.2]
         */
        readonly isLocked: boolean;
        /**
         *
         * Gets the name of the section. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly name: string;
        /**
         *
         * The web url of the page. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly webUrl: string;
        /**
         *
         * Adds a new page to the end of the section.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param title - The title of the new page.
         */
        addPage(title: string): OneNote.Page;
        /**
         *
         * Copies this section to specified notebook.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param destinationNotebook - The notebook to copy this section to.
         */
        copyToNotebook(destinationNotebook: OneNote.Notebook): OneNote.Section;
        /**
         *
         * Copies this section to specified section group.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param destinationSectionGroup - The section group to copy this section to.
         */
        copyToSectionGroup(destinationSectionGroup: OneNote.SectionGroup): OneNote.Section;
        /**
         *
         * Gets the REST API ID.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getRestApiId(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Inserts a new section before or after the current section.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param location - The location of the new section relative to the current section.
         * @param title - The name of the new section.
         */
        insertSectionAsSibling(location: OneNote.InsertLocation, title: string): OneNote.Section;
        /**
         *
         * Inserts a new section before or after the current section.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param locationString - The location of the new section relative to the current section.
         * @param title - The name of the new section.
         */
        insertSectionAsSibling(locationString: "Before" | "After", title: string): OneNote.Section;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.Section` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.Section` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.Section` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.SectionLoadOptions): OneNote.Section;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.Section;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): OneNote.Section;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.Section;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.Section;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OneNote.Section object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.SectionData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): OneNote.Interfaces.SectionData;
    }
    /**
     *
     * Represents a collection of sections.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class SectionCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: OneNote.Section[];
        /**
         *
         * Returns the number of sections in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        /**
         *
         * Gets the collection of sections with the specified name.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param name - The name of the section.
         */
        getByName(name: string): OneNote.SectionCollection;
        /**
         *
         * Gets a section by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - The ID of the section, or the index location of the section in the collection.
         */
        getItem(index: number | string): OneNote.Section;
        /**
         *
         * Gets a section on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.Section;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.SectionCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.SectionCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.SectionCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.SectionCollectionLoadOptions & OneNote.Interfaces.CollectionLoadOptions): OneNote.SectionCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.SectionCollection;
        load(option?: OfficeExtension.LoadOption): OneNote.SectionCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.SectionCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.SectionCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `OneNote.SectionCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.SectionCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): OneNote.Interfaces.SectionCollectionData;
    }
    /**
     *
     * Represents a OneNote page.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class Page extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * The collection of PageContent objects on the page. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly contents: OneNote.PageContentCollection;
        /**
         *
         * Text interpretation for the ink on the page. Returns null if there is no ink analysis information. Read only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly inkAnalysisOrNull: OneNote.InkAnalysis;
        /**
         *
         * Gets the section that contains the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentSection: OneNote.Section;
        /**
         *
         * Gets the ClassNotebookPageSource to the page.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly classNotebookPageSource: string;
        /**
         *
         * The client url of the page. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly clientUrl: string;
        /**
         *
         * Gets the ID of the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * Gets or sets the indentation level of the page.
         *
         * [Api set: OneNoteApi 1.1]
         */
        pageLevel: number;
        /**
         *
         * Gets or sets the title of the page.
         *
         * [Api set: OneNoteApi 1.1]
         */
        title: string;
        /**
         *
         * The web url of the page. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly webUrl: string;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: OneNote.Page): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.PageUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: OneNote.Page): void;
        /**
         *
         * Adds an Outline to the page at the specified position.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param left - The left position of the top, left corner of the Outline.
         * @param top - The top position of the top, left corner of the Outline.
         * @param html - An HTML string that describes the visual presentation of the Outline. See {@link https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-page-content#supported-html | Supported HTML} for the OneNote add-ins JavaScript API.
         */
        addOutline(left: number, top: number, html: string): OneNote.Outline;
        /**
         *
         * Return a json string with node id and content in html format.
         *
         * [Api set: OneNoteApi 1.1]
         */
        analyzePage(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Inserts a new page with translated content.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param translatedContent - Translated content of the page
         */
        applyTranslation(translatedContent: string): void;
        /**
         *
         * Copies this page to specified section.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param destinationSection - The section to copy this page to.
         */
        copyToSection(destinationSection: OneNote.Section): OneNote.Page;
        /**
         *
         * Copies this page to specified section and sets ClassNotebookPageSource.
         *
         * [Api set: OneNoteApi 1.1]
         */
        copyToSectionAndSetClassNotebookPageSource(destinationSection: OneNote.Section): OneNote.Page;
        /**
         *
         * Gets the REST API ID.
         *
         * [Api set: OneNoteApi 1.1]
         */
        getRestApiId(): OfficeExtension.ClientResult<string>;
        /**
         *
         * Does the page has content title.
         *
         * [Api set: OneNoteApi 1.1]
         */
        hasTitleContent(): OfficeExtension.ClientResult<boolean>;
        /**
         *
         * Inserts a new page before or after the current page.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param location - The location of the new page relative to the current page.
         * @param title - The title of the new page.
         */
        insertPageAsSibling(location: OneNote.InsertLocation, title: string): OneNote.Page;
        /**
         *
         * Inserts a new page before or after the current page.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param locationString - The location of the new page relative to the current page.
         * @param title - The title of the new page.
         */
        insertPageAsSibling(locationString: "Before" | "After", title: string): OneNote.Page;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.Page` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.Page` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.Page` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.PageLoadOptions): OneNote.Page;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.Page;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): OneNote.Page;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.Page;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.Page;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OneNote.Page object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.PageData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): OneNote.Interfaces.PageData;
    }
    /**
     *
     * Represents a collection of pages.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class PageCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: OneNote.Page[];
        /**
         *
         * Returns the number of pages in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        /**
         *
         * Gets the collection of pages with the specified title.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param title - The title of the page.
         */
        getByTitle(title: string): OneNote.PageCollection;
        /**
         *
         * Gets a page by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - The ID of the page, or the index location of the page in the collection.
         */
        getItem(index: number | string): OneNote.Page;
        /**
         *
         * Gets a page on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.Page;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.PageCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.PageCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.PageCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.PageCollectionLoadOptions & OneNote.Interfaces.CollectionLoadOptions): OneNote.PageCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.PageCollection;
        load(option?: OfficeExtension.LoadOption): OneNote.PageCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.PageCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.PageCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `OneNote.PageCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.PageCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): OneNote.Interfaces.PageCollectionData;
    }
    /**
     *
     * Represents a region on a page that contains top-level content types such as Outline or Image. A PageContent object can be assigned an XY position.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class PageContent extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the Image in the PageContent object. Throws an exception if PageContentType is not Image.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly image: OneNote.Image;
        /**
         *
         * Gets the ink in the PageContent object. Throws an exception if PageContentType is not Ink.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly ink: OneNote.FloatingInk;
        /**
         *
         * Gets the Outline in the PageContent object. Throws an exception if PageContentType is not Outline.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly outline: OneNote.Outline;
        /**
         *
         * Gets the page that contains the PageContent object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentPage: OneNote.Page;
        /**
         *
         * Gets the ID of the PageContent object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * Gets or sets the left (X-axis) position of the PageContent object.
         *
         * [Api set: OneNoteApi 1.1]
         */
        left: number;
        /**
         *
         * Gets or sets the top (Y-axis) position of the PageContent object.
         *
         * [Api set: OneNoteApi 1.1]
         */
        top: number;
        /**
         *
         * Gets the type of the PageContent object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly type: OneNote.PageContentType | "Outline" | "Image" | "Ink" | "Other";
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: OneNote.PageContent): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.PageContentUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: OneNote.PageContent): void;
        /**
         *
         * Deletes the PageContent object.
         *
         * [Api set: OneNoteApi 1.1]
         */
        delete(): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.PageContent` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.PageContent` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.PageContent` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.PageContentLoadOptions): OneNote.PageContent;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.PageContent;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): OneNote.PageContent;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.PageContent;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.PageContent;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OneNote.PageContent object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.PageContentData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): OneNote.Interfaces.PageContentData;
    }
    /**
     *
     * Represents the contents of a page, as a collection of PageContent objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class PageContentCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: OneNote.PageContent[];
        /**
         *
         * Returns the number of page contents in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        /**
         *
         * Gets a PageContent object by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - The ID of the PageContent object, or the index location of the PageContent object in the collection.
         */
        getItem(index: number | string): OneNote.PageContent;
        /**
         *
         * Gets a page content on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.PageContent;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.PageContentCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.PageContentCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.PageContentCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.PageContentCollectionLoadOptions & OneNote.Interfaces.CollectionLoadOptions): OneNote.PageContentCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.PageContentCollection;
        load(option?: OfficeExtension.LoadOption): OneNote.PageContentCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.PageContentCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.PageContentCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `OneNote.PageContentCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.PageContentCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): OneNote.Interfaces.PageContentCollectionData;
    }
    /**
     *
     * Represents a container for Paragraph objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class Outline extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the PageContent object that contains the Outline. This object defines the position of the Outline on the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly pageContent: OneNote.PageContent;
        /**
         *
         * Gets the collection of Paragraph objects in the Outline. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly paragraphs: OneNote.ParagraphCollection;
        /**
         *
         * Gets the ID of the Outline object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * Adds the specified HTML to the bottom of the Outline.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param html - The HTML string to append. See {@link https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-page-content#supported-html | Supported HTML} for the OneNote add-ins JavaScript API.
         */
        appendHtml(html: string): void;
        /**
         *
         * Adds the specified image to the bottom of the Outline.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param base64EncodedImage - HTML string to append.
         * @param width - Optional. Width in the unit of Points. The default value is null and image width will be respected.
         * @param height - Optional. Height in the unit of Points. The default value is null and image height will be respected.
         */
        appendImage(base64EncodedImage: string, width: number, height: number): OneNote.Image;
        /**
         *
         * Adds the specified text to the bottom of the Outline.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param paragraphText - HTML string to append.
         */
        appendRichText(paragraphText: string): OneNote.RichText;
        /**
         *
         * Adds a table with the specified number of rows and columns to the bottom of the outline.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param rowCount - Required. The number of rows in the table.
         * @param columnCount - Required. The number of columns in the table.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        appendTable(rowCount: number, columnCount: number, values?: string[][]): OneNote.Table;
        /**
         *
         * Check if the outline is title outline.
         *
         * [Api set: OneNoteApi 1.1]
         */
        isTitle(): OfficeExtension.ClientResult<boolean>;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.Outline` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.Outline` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.Outline` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.OutlineLoadOptions): OneNote.Outline;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.Outline;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): OneNote.Outline;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.Outline;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.Outline;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OneNote.Outline object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.OutlineData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): OneNote.Interfaces.OutlineData;
    }
    /**
     *
     * A container for the visible content on a page. A Paragraph can contain any one ParagraphType type of content.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class Paragraph extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the Image object in the Paragraph. Throws an exception if ParagraphType is not Image. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly image: OneNote.Image;
        /**
         *
         * Gets the Ink collection in the Paragraph. Throws an exception if ParagraphType is not Ink. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly inkWords: OneNote.InkWordCollection;
        /**
         *
         * Gets the Outline object that contains the Paragraph. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly outline: OneNote.Outline;
        /**
         *
         * The collection of paragraphs under this paragraph. Read only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly paragraphs: OneNote.ParagraphCollection;
        /**
         *
         * Gets the parent paragraph object. Throws if a parent paragraph does not exist. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentParagraph: OneNote.Paragraph;
        /**
         *
         * Gets the parent paragraph object. Returns null if a parent paragraph does not exist. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentParagraphOrNull: OneNote.Paragraph;
        /**
         *
         * Gets the TableCell object that contains the Paragraph if one exists. If parent is not a TableCell, throws ItemNotFound. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentTableCell: OneNote.TableCell;
        /**
         *
         * Gets the TableCell object that contains the Paragraph if one exists. If parent is not a TableCell, returns null. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentTableCellOrNull: OneNote.TableCell;
        /**
         *
         * Gets the RichText object in the Paragraph. Throws an exception if ParagraphType is not RichText. Read-only
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly richText: OneNote.RichText;
        /**
         *
         * Gets the Table object in the Paragraph. Throws an exception if ParagraphType is not Table. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly table: OneNote.Table;
        /**
         *
         * Gets the ID of the Paragraph object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * Gets the type of the Paragraph object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly type: OneNote.ParagraphType | "RichText" | "Image" | "Table" | "Ink" | "Other";
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: OneNote.Paragraph): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ParagraphUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: OneNote.Paragraph): void;
        /**
         *
         * Add NoteTag to the paragraph.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param type - The type of the NoteTag.
         * @param status - The status of the NoteTag.
         */
        addNoteTag(type: OneNote.NoteTagType, status: OneNote.NoteTagStatus): OneNote.NoteTag;
        /**
         *
         * Add NoteTag to the paragraph.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param typeString - The type of the NoteTag.
         * @param status - The status of the NoteTag.
         */
        addNoteTag(typeString: "Unknown" | "ToDo" | "Important" | "Question" | "Contact" | "Address" | "PhoneNumber" | "Website" | "Idea" | "Critical" | "ToDoPriority1" | "ToDoPriority2", status: "Unknown" | "Normal" | "Completed" | "Disabled" | "OutlookTask" | "TaskNotSyncedYet" | "TaskRemoved"): OneNote.NoteTag;
        /**
         *
         * Deletes the paragraph
         *
         * [Api set: OneNoteApi 1.1]
         */
        delete(): void;
        /**
         *
         * Get list information of paragraph
         *
         * [Api set: OneNoteApi 1.1]
         */
        getParagraphInfo(): OfficeExtension.ClientResult<OneNote.ParagraphInfo>;
        /**
         *
         * Inserts the specified HTML content
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param insertLocation - The location of new contents relative to the current Paragraph.
         * @param html - An HTML string that describes the visual presentation of the content. See {@link https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-page-content#supported-html | Supported HTML} for the OneNote add-ins JavaScript API.
         */
        insertHtmlAsSibling(insertLocation: OneNote.InsertLocation, html: string): void;
        /**
         *
         * Inserts the specified HTML content
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param insertLocationString - The location of new contents relative to the current Paragraph.
         * @param html - An HTML string that describes the visual presentation of the content. See {@link https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-page-content#supported-html | Supported HTML} for the OneNote add-ins JavaScript API.
         */
        insertHtmlAsSibling(insertLocationString: "Before" | "After", html: string): void;
        /**
         *
         * Inserts the image at the specified insert location..
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param insertLocation - The location of the table relative to the current Paragraph.
         * @param base64EncodedImage - HTML string to append.
         * @param width - Optional. Width in the unit of Points. The default value is null and image width will be respected.
         * @param height - Optional. Height in the unit of Points. The default value is null and image height will be respected.
         */
        insertImageAsSibling(insertLocation: OneNote.InsertLocation, base64EncodedImage: string, width: number, height: number): OneNote.Image;
        /**
         *
         * Inserts the image at the specified insert location..
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param insertLocationString - The location of the table relative to the current Paragraph.
         * @param base64EncodedImage - HTML string to append.
         * @param width - Optional. Width in the unit of Points. The default value is null and image width will be respected.
         * @param height - Optional. Height in the unit of Points. The default value is null and image height will be respected.
         */
        insertImageAsSibling(insertLocationString: "Before" | "After", base64EncodedImage: string, width: number, height: number): OneNote.Image;
        /**
         *
         * Inserts the paragraph text at the specifiec insert location.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param insertLocation - The location of the table relative to the current Paragraph.
         * @param paragraphText - HTML string to append.
         */
        insertRichTextAsSibling(insertLocation: OneNote.InsertLocation, paragraphText: string): OneNote.RichText;
        /**
         *
         * Inserts the paragraph text at the specifiec insert location.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param insertLocationString - The location of the table relative to the current Paragraph.
         * @param paragraphText - HTML string to append.
         */
        insertRichTextAsSibling(insertLocationString: "Before" | "After", paragraphText: string): OneNote.RichText;
        /**
         *
         * Adds a table with the specified number of rows and columns before or after the current paragraph.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param insertLocation - The location of the table relative to the current Paragraph.
         * @param rowCount - The number of rows in the table.
         * @param columnCount - The number of columns in the table.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        insertTableAsSibling(insertLocation: OneNote.InsertLocation, rowCount: number, columnCount: number, values?: string[][]): OneNote.Table;
        /**
         *
         * Adds a table with the specified number of rows and columns before or after the current paragraph.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param insertLocationString - The location of the table relative to the current Paragraph.
         * @param rowCount - The number of rows in the table.
         * @param columnCount - The number of columns in the table.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        insertTableAsSibling(insertLocationString: "Before" | "After", rowCount: number, columnCount: number, values?: string[][]): OneNote.Table;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.Paragraph` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.Paragraph` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.Paragraph` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.ParagraphLoadOptions): OneNote.Paragraph;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.Paragraph;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): OneNote.Paragraph;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.Paragraph;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.Paragraph;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OneNote.Paragraph object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.ParagraphData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): OneNote.Interfaces.ParagraphData;
    }
    /**
     *
     * Represents a collection of Paragraph objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class ParagraphCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: OneNote.Paragraph[];
        /**
         *
         * Returns the number of paragraphs in the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        /**
         *
         * Gets a Paragraph object by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - The ID of the Paragraph object, or the index location of the Paragraph object in the collection.
         */
        getItem(index: number | string): OneNote.Paragraph;
        /**
         *
         * Gets a paragraph on its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.Paragraph;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.ParagraphCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.ParagraphCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.ParagraphCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.ParagraphCollectionLoadOptions & OneNote.Interfaces.CollectionLoadOptions): OneNote.ParagraphCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.ParagraphCollection;
        load(option?: OfficeExtension.LoadOption): OneNote.ParagraphCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.ParagraphCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.ParagraphCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `OneNote.ParagraphCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.ParagraphCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): OneNote.Interfaces.ParagraphCollectionData;
    }
    /**
     *
     * A container for the NoteTag in a paragraph.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class NoteTag extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the Id of the NoteTag object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * Gets the status of the NoteTag object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly status: OneNote.NoteTagStatus | "Unknown" | "Normal" | "Completed" | "Disabled" | "OutlookTask" | "TaskNotSyncedYet" | "TaskRemoved";
        /**
         *
         * Gets the type of the NoteTag object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly type: OneNote.NoteTagType | "Unknown" | "ToDo" | "Important" | "Question" | "Contact" | "Address" | "PhoneNumber" | "Website" | "Idea" | "Critical" | "ToDoPriority1" | "ToDoPriority2";
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.NoteTag` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.NoteTag` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.NoteTag` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.NoteTagLoadOptions): OneNote.NoteTag;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.NoteTag;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): OneNote.NoteTag;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.NoteTag;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.NoteTag;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OneNote.NoteTag object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.NoteTagData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): OneNote.Interfaces.NoteTagData;
    }
    /**
     *
     * Represents a RichText object in a Paragraph.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class RichText extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the Paragraph object that contains the RichText object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly paragraph: OneNote.Paragraph;
        /**
         *
         * Gets the ID of the RichText object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * The language id of the text. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly languageId: string;
        /**
         *
         * Gets the text content of the RichText object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly text: string;
        /**
         *
         * Get the HTML of the rich text
         *
         * [Api set: OneNoteApi 1.1]
         * @returns The html of the rich text
         */
        getHtml(): OfficeExtension.ClientResult<string>;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.RichText` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.RichText` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.RichText` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.RichTextLoadOptions): OneNote.RichText;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.RichText;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): OneNote.RichText;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.RichText;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.RichText;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OneNote.RichText object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.RichTextData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): OneNote.Interfaces.RichTextData;
    }
    /**
     *
     * Represents an Image. An Image can be a direct child of a PageContent object or a Paragraph object.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class Image extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the PageContent object that contains the Image. Throws if the Image is not a direct child of a PageContent. This object defines the position of the Image on the page. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly pageContent: OneNote.PageContent;
        /**
         *
         * Gets the Paragraph object that contains the Image. Throws if the Image is not a direct child of a Paragraph. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly paragraph: OneNote.Paragraph;
        /**
         *
         * Gets or sets the description of the Image.
         *
         * [Api set: OneNoteApi 1.1]
         */
        description: string;
        /**
         *
         * Gets or sets the height of the Image layout.
         *
         * [Api set: OneNoteApi 1.1]
         */
        height: number;
        /**
         *
         * Gets or sets the hyperlink of the Image.
         *
         * [Api set: OneNoteApi 1.1]
         */
        hyperlink: string;
        /**
         *
         * Gets the ID of the Image object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * Gets the data obtained by OCR (Optical Character Recognition) of this Image, such as OCR text and language.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly ocrData: OneNote.ImageOcrData;
        /**
         *
         * Gets or sets the width of the Image layout.
         *
         * [Api set: OneNoteApi 1.1]
         */
        width: number;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: OneNote.Image): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ImageUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: OneNote.Image): void;
        /**
         *
         * Gets the base64-encoded binary representation of the Image.
            Example: data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADIA...
         *
         * [Api set: OneNoteApi 1.1]
         */
        getBase64Image(): OfficeExtension.ClientResult<string>;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.Image` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.Image` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.Image` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.ImageLoadOptions): OneNote.Image;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.Image;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): OneNote.Image;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.Image;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.Image;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OneNote.Image object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.ImageData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): OneNote.Interfaces.ImageData;
    }
    /**
     *
     * Represents a table in a OneNote page.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class Table extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the Paragraph object that contains the Table object. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly paragraph: OneNote.Paragraph;
        /**
         *
         * Gets all of the table rows. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly rows: OneNote.TableRowCollection;
        /**
         *
         * Gets or sets whether the borders are visible or not. True if they are visible, false if they are hidden.
         *
         * [Api set: OneNoteApi 1.1]
         */
        borderVisible: boolean;
        /**
         *
         * Gets the number of columns in the table.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly columnCount: number;
        /**
         *
         * Gets the ID of the table. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * Gets the number of rows in the table.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly rowCount: number;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: OneNote.Table): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.TableUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: OneNote.Table): void;
        /**
         *
         * Adds a column to the end of the table. Values, if specified, are set in the new column. Otherwise the column is empty.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param values - Optional. Strings to insert in the new column, specified as an array. Must not have more values than rows in the table.
         */
        appendColumn(values?: string[]): void;
        /**
         *
         * Adds a row to the end of the table. Values, if specified, are set in the new row. Otherwise the row is empty.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param values - Optional. Strings to insert in the new row, specified as an array. Must not have more values than columns in the table.
         */
        appendRow(values?: string[]): OneNote.TableRow;
        /**
         *
         * Clears the contents of the table.
         *
         * [Api set: OneNoteApi 1.1]
         */
        clear(): void;
        /**
         *
         * Gets the table cell at a specified row and column.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param rowIndex - The index of the row.
         * @param cellIndex - The index of the cell in the row.
         */
        getCell(rowIndex: number, cellIndex: number): OneNote.TableCell;
        /**
         *
         * Inserts a column at the given index in the table. Values, if specified, are set in the new column. Otherwise the column is empty.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - Index where the column will be inserted in the table.
         * @param values - Optional. Strings to insert in the new column, specified as an array. Must not have more values than rows in the table.
         */
        insertColumn(index: number, values?: string[]): void;
        /**
         *
         * Inserts a row at the given index in the table. Values, if specified, are set in the new row. Otherwise the row is empty.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - Index where the row will be inserted in the table.
         * @param values - Optional. Strings to insert in the new row, specified as an array. Must not have more values than columns in the table.
         */
        insertRow(index: number, values?: string[]): OneNote.TableRow;
        /**
         *
         * Sets the shading color of all cells in the table.
            The color code to set the cells to.
         *
         * [Api set: OneNoteApi 1.1]
         */
        setShadingColor(colorCode: string): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.Table` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.Table` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.Table` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.TableLoadOptions): OneNote.Table;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.Table;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): OneNote.Table;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.Table;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.Table;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OneNote.Table object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.TableData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): OneNote.Interfaces.TableData;
    }
    /**
     *
     * Represents a row in a table.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class TableRow extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the cells in the row. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly cells: OneNote.TableCellCollection;
        /**
         *
         * Gets the parent table. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentTable: OneNote.Table;
        /**
         *
         * Gets the number of cells in the row. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly cellCount: number;
        /**
         *
         * Gets the ID of the row. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * Gets the index of the row in its parent table. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly rowIndex: number;
        /**
         *
         * Clears the contents of the row.
         *
         * [Api set: OneNoteApi 1.1]
         */
        clear(): void;
        /**
         *
         * Inserts a row before or after the current row.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param insertLocation - Where the new rows should be inserted relative to the current row.
         * @param values - Strings to insert in the new row, specified as an array. Must not have more cells than in the current row. Optional.
         */
        insertRowAsSibling(insertLocation: OneNote.InsertLocation, values?: string[]): OneNote.TableRow;
        /**
         *
         * Inserts a row before or after the current row.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param insertLocationString - Where the new rows should be inserted relative to the current row.
         * @param values - Strings to insert in the new row, specified as an array. Must not have more cells than in the current row. Optional.
         */
        insertRowAsSibling(insertLocationString: "Before" | "After", values?: string[]): OneNote.TableRow;
        /**
         *
         * Sets the shading color of all cells in the row.
            The color code to set the cells to.
         *
         * [Api set: OneNoteApi 1.1]
         */
        setShadingColor(colorCode: string): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.TableRow` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.TableRow` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.TableRow` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.TableRowLoadOptions): OneNote.TableRow;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.TableRow;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): OneNote.TableRow;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.TableRow;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.TableRow;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OneNote.TableRow object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.TableRowData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): OneNote.Interfaces.TableRowData;
    }
    /**
     *
     * Contains a collection of TableRow objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class TableRowCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: OneNote.TableRow[];
        /**
         *
         * Returns the number of table rows in this collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        /**
         *
         * Gets a table row object by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - A number that identifies the index location of a table row object.
         */
        getItem(index: number | string): OneNote.TableRow;
        /**
         *
         * Gets a table row at its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.TableRow;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.TableRowCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.TableRowCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.TableRowCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.TableRowCollectionLoadOptions & OneNote.Interfaces.CollectionLoadOptions): OneNote.TableRowCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.TableRowCollection;
        load(option?: OfficeExtension.LoadOption): OneNote.TableRowCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.TableRowCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.TableRowCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `OneNote.TableRowCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.TableRowCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): OneNote.Interfaces.TableRowCollectionData;
    }
    /**
     *
     * Represents a cell in a OneNote table.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class TableCell extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the collection of Paragraph objects in the TableCell. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly paragraphs: OneNote.ParagraphCollection;
        /**
         *
         * Gets the parent row of the cell. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly parentRow: OneNote.TableRow;
        /**
         *
         * Gets the index of the cell in its row. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly cellIndex: number;
        /**
         *
         * Gets the ID of the cell. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly id: string;
        /**
         *
         * Gets the index of the cell's row in the table. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly rowIndex: number;
        /**
         *
         * Gets and sets the shading color of the cell
         *
         * [Api set: OneNoteApi 1.1]
         */
        shadingColor: string;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: OneNote.TableCell): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.TableCellUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: OneNote.TableCell): void;
        /**
         *
         * Adds the specified HTML to the bottom of the TableCell.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param html - The HTML string to append. See {@link https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-page-content#supported-html | Supported HTML} for the OneNote add-ins JavaScript API.
         */
        appendHtml(html: string): void;
        /**
         *
         * Adds the specified image to table cell.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param base64EncodedImage - HTML string to append.
         * @param width - Optional. Width in the unit of Points. The default value is null and image width will be respected.
         * @param height - Optional. Height in the unit of Points. The default value is null and image height will be respected.
         */
        appendImage(base64EncodedImage: string, width: number, height: number): OneNote.Image;
        /**
         *
         * Adds the specified text to table cell.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param paragraphText - HTML string to append.
         */
        appendRichText(paragraphText: string): OneNote.RichText;
        /**
         *
         * Adds a table with the specified number of rows and columns to table cell.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param rowCount - Required. The number of rows in the table.
         * @param columnCount - Required. The number of columns in the table.
         * @param values - Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
         */
        appendTable(rowCount: number, columnCount: number, values?: string[][]): OneNote.Table;
        /**
         *
         * Clears the contents of the cell.
         *
         * [Api set: OneNoteApi 1.1]
         */
        clear(): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.TableCell` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.TableCell` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.TableCell` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.TableCellLoadOptions): OneNote.TableCell;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.TableCell;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): OneNote.TableCell;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.TableCell;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.TableCell;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original OneNote.TableCell object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.TableCellData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): OneNote.Interfaces.TableCellData;
    }
    /**
     *
     * Contains a collection of TableCell objects.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export class TableCellCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: OneNote.TableCell[];
        /**
         *
         * Returns the number of tablecells in this collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         */
        readonly count: number;
        /**
         *
         * Gets a table cell object by ID or by its index in the collection. Read-only.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - A number that identifies the index location of a table cell object.
         */
        getItem(index: number | string): OneNote.TableCell;
        /**
         *
         * Gets a tablecell at its position in the collection.
         *
         * [Api set: OneNoteApi 1.1]
         *
         * @param index - Index value of the object to be retrieved. Zero-indexed.
         */
        getItemAt(index: number): OneNote.TableCell;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): OneNote.TableCellCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): OneNote.TableCellCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): OneNote.TableCellCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: OneNote.Interfaces.TableCellCollectionLoadOptions & OneNote.Interfaces.CollectionLoadOptions): OneNote.TableCellCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): OneNote.TableCellCollection;
        load(option?: OfficeExtension.LoadOption): OneNote.TableCellCollection;
        /**
         * Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you are using this object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created.
         */
        track(): OneNote.TableCellCollection;
        /**
         * Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect.
         */
        untrack(): OneNote.TableCellCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `OneNote.TableCellCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `OneNote.Interfaces.TableCellCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): OneNote.Interfaces.TableCellCollectionData;
    }
    /**
     *
     * Represents data obtained by OCR (optical character recognition) of an image.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export interface ImageOcrData {
        /**
         *
         * Represents the OCR language, with values such as EN-US
         *
         * [Api set: OneNoteApi 1.1]
         */
        ocrLanguageId: string;
        /**
         *
         * Represents the text obtained by OCR of the image
         *
         * [Api set: OneNoteApi 1.1]
         */
        ocrText: string;
    }
    /**
     *
     * Weak reference to an ink stroke object and its content parent.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export interface InkStrokePointer {
        /**
         *
         * Represents the id of the page content object corresponding to this stroke
         *
         * [Api set: OneNoteApi 1.1]
         */
        contentId: string;
        /**
         *
         * Represents the id of the ink stroke
         *
         * [Api set: OneNoteApi 1.1]
         */
        inkStrokeId: string;
    }
    /**
     *
     * List information for paragraph.
     *
     * [Api set: OneNoteApi 1.1]
     */
    export interface ParagraphInfo {
        /**
         *
         * //
            Bullet list type of paragraph
         *
         * [Api set: OneNoteApi 1.1]
         */
        bulletType: string;
        /**
         *
         * //
            Index of paragraph in list
         *
         * [Api set: OneNoteApi 1.1]
         */
        index: number;
        /**
         *
         * //
            Type of list in paragraph
         *
         * [Api set: OneNoteApi 1.1]
         */
        listType: OneNote.ListType | "None" | "Number" | "Bullet";
        /**
         *
         * //
            number list type of paragraph
         *
         * [Api set: OneNoteApi 1.1]
         */
        numberType: OneNote.NumberType | "None" | "Arabic" | "UCRoman" | "LCRoman" | "UCLetter" | "LCLetter" | "Ordinal" | "Cardtext" | "Ordtext" | "Hex" | "ChiManSty" | "DbNum1" | "DbNum2" | "Aiueo" | "Iroha" | "DbChar" | "SbChar" | "DbNum3" | "DbNum4" | "Circlenum" | "DArabic" | "DAiueo" | "DIroha" | "ArabicLZ" | "Bullet" | "Ganada" | "Chosung" | "GB1" | "GB2" | "GB3" | "GB4" | "Zodiac1" | "Zodiac2" | "Zodiac3" | "TpeDbNum1" | "TpeDbNum2" | "TpeDbNum3" | "TpeDbNum4" | "ChnDbNum1" | "ChnDbNum2" | "ChnDbNum3" | "ChnDbNum4" | "KorDbNum1" | "KorDbNum2" | "KorDbNum3" | "KorDbNum4" | "Hebrew1" | "Arabic1" | "Hebrew2" | "Arabic2" | "Hindi1" | "Hindi2" | "Hindi3" | "Thai1" | "Thai2" | "NumInDash" | "LCRus" | "UCRus" | "LCGreek" | "UCGreek" | "Lim" | "Custom";
    }
    /**
     * [Api set: OneNoteApi 1.1]
     */
    enum InsertLocation {
        before = "Before",
        after = "After",
    }
    /**
     * [Api set: OneNoteApi 1.1]
     */
    enum PageContentType {
        outline = "Outline",
        image = "Image",
        ink = "Ink",
        other = "Other",
    }
    /**
     * [Api set: OneNoteApi 1.1]
     */
    enum ParagraphType {
        richText = "RichText",
        image = "Image",
        table = "Table",
        ink = "Ink",
        other = "Other",
    }
    /**
     * [Api set: OneNoteApi 1.1]
     */
    enum NoteTagType {
        unknown = "Unknown",
        toDo = "ToDo",
        important = "Important",
        question = "Question",
        contact = "Contact",
        address = "Address",
        phoneNumber = "PhoneNumber",
        website = "Website",
        idea = "Idea",
        critical = "Critical",
        toDoPriority1 = "ToDoPriority1",
        toDoPriority2 = "ToDoPriority2",
    }
    /**
     * [Api set: OneNoteApi 1.1]
     */
    enum NoteTagStatus {
        unknown = "Unknown",
        normal = "Normal",
        completed = "Completed",
        disabled = "Disabled",
        outlookTask = "OutlookTask",
        taskNotSyncedYet = "TaskNotSyncedYet",
        taskRemoved = "TaskRemoved",
    }
    /**
     * [Api set: OneNoteApi 1.1]
     */
    enum ListType {
        none = "None",
        number = "Number",
        bullet = "Bullet",
    }
    /**
     * [Api set: OneNoteApi 1.1]
     */
    enum NumberType {
        none = "None",
        arabic = "Arabic",
        ucroman = "UCRoman",
        lcroman = "LCRoman",
        ucletter = "UCLetter",
        lcletter = "LCLetter",
        ordinal = "Ordinal",
        cardtext = "Cardtext",
        ordtext = "Ordtext",
        hex = "Hex",
        chiManSty = "ChiManSty",
        dbNum1 = "DbNum1",
        dbNum2 = "DbNum2",
        aiueo = "Aiueo",
        iroha = "Iroha",
        dbChar = "DbChar",
        sbChar = "SbChar",
        dbNum3 = "DbNum3",
        dbNum4 = "DbNum4",
        circlenum = "Circlenum",
        darabic = "DArabic",
        daiueo = "DAiueo",
        diroha = "DIroha",
        arabicLZ = "ArabicLZ",
        bullet = "Bullet",
        ganada = "Ganada",
        chosung = "Chosung",
        gb1 = "GB1",
        gb2 = "GB2",
        gb3 = "GB3",
        gb4 = "GB4",
        zodiac1 = "Zodiac1",
        zodiac2 = "Zodiac2",
        zodiac3 = "Zodiac3",
        tpeDbNum1 = "TpeDbNum1",
        tpeDbNum2 = "TpeDbNum2",
        tpeDbNum3 = "TpeDbNum3",
        tpeDbNum4 = "TpeDbNum4",
        chnDbNum1 = "ChnDbNum1",
        chnDbNum2 = "ChnDbNum2",
        chnDbNum3 = "ChnDbNum3",
        chnDbNum4 = "ChnDbNum4",
        korDbNum1 = "KorDbNum1",
        korDbNum2 = "KorDbNum2",
        korDbNum3 = "KorDbNum3",
        korDbNum4 = "KorDbNum4",
        hebrew1 = "Hebrew1",
        arabic1 = "Arabic1",
        hebrew2 = "Hebrew2",
        arabic2 = "Arabic2",
        hindi1 = "Hindi1",
        hindi2 = "Hindi2",
        hindi3 = "Hindi3",
        thai1 = "Thai1",
        thai2 = "Thai2",
        numInDash = "NumInDash",
        lcrus = "LCRus",
        ucrus = "UCRus",
        lcgreek = "LCGreek",
        ucgreek = "UCGreek",
        lim = "Lim",
        custom = "Custom",
    }
    enum ErrorCodes {
        generalException = "GeneralException",
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
        /** An interface for updating data on the InkAnalysis object, for use in "inkAnalysis.set({ ... })". */
        export interface InkAnalysisUpdateData {
            /**
            *
            * Gets the parent page object.
            *
            * [Api set: OneNoteApi 1.1]
            */
            page?: OneNote.Interfaces.PageUpdateData;
        }
        /** An interface for updating data on the InkAnalysisParagraph object, for use in "inkAnalysisParagraph.set({ ... })". */
        export interface InkAnalysisParagraphUpdateData {
            /**
            *
            * Reference to the parent InkAnalysisPage.
            *
            * [Api set: OneNoteApi 1.1]
            */
            inkAnalysis?: OneNote.Interfaces.InkAnalysisUpdateData;
        }
        /** An interface for updating data on the InkAnalysisParagraphCollection object, for use in "inkAnalysisParagraphCollection.set({ ... })". */
        export interface InkAnalysisParagraphCollectionUpdateData {
            items?: OneNote.Interfaces.InkAnalysisParagraphData[];
        }
        /** An interface for updating data on the InkAnalysisLine object, for use in "inkAnalysisLine.set({ ... })". */
        export interface InkAnalysisLineUpdateData {
            /**
            *
            * Reference to the parent InkAnalysisParagraph.
            *
            * [Api set: OneNoteApi 1.1]
            */
            paragraph?: OneNote.Interfaces.InkAnalysisParagraphUpdateData;
        }
        /** An interface for updating data on the InkAnalysisLineCollection object, for use in "inkAnalysisLineCollection.set({ ... })". */
        export interface InkAnalysisLineCollectionUpdateData {
            items?: OneNote.Interfaces.InkAnalysisLineData[];
        }
        /** An interface for updating data on the InkAnalysisWord object, for use in "inkAnalysisWord.set({ ... })". */
        export interface InkAnalysisWordUpdateData {
            /**
            *
            * Reference to the parent InkAnalysisLine.
            *
            * [Api set: OneNoteApi 1.1]
            */
            line?: OneNote.Interfaces.InkAnalysisLineUpdateData;
        }
        /** An interface for updating data on the InkAnalysisWordCollection object, for use in "inkAnalysisWordCollection.set({ ... })". */
        export interface InkAnalysisWordCollectionUpdateData {
            items?: OneNote.Interfaces.InkAnalysisWordData[];
        }
        /** An interface for updating data on the InkStrokeCollection object, for use in "inkStrokeCollection.set({ ... })". */
        export interface InkStrokeCollectionUpdateData {
            items?: OneNote.Interfaces.InkStrokeData[];
        }
        /** An interface for updating data on the InkWordCollection object, for use in "inkWordCollection.set({ ... })". */
        export interface InkWordCollectionUpdateData {
            items?: OneNote.Interfaces.InkWordData[];
        }
        /** An interface for updating data on the NotebookCollection object, for use in "notebookCollection.set({ ... })". */
        export interface NotebookCollectionUpdateData {
            items?: OneNote.Interfaces.NotebookData[];
        }
        /** An interface for updating data on the SectionGroupCollection object, for use in "sectionGroupCollection.set({ ... })". */
        export interface SectionGroupCollectionUpdateData {
            items?: OneNote.Interfaces.SectionGroupData[];
        }
        /** An interface for updating data on the SectionCollection object, for use in "sectionCollection.set({ ... })". */
        export interface SectionCollectionUpdateData {
            items?: OneNote.Interfaces.SectionData[];
        }
        /** An interface for updating data on the Page object, for use in "page.set({ ... })". */
        export interface PageUpdateData {
            /**
            *
            * Text interpretation for the ink on the page. Returns null if there is no ink analysis information. Read only.
            *
            * [Api set: OneNoteApi 1.1]
            */
            inkAnalysisOrNull?: OneNote.Interfaces.InkAnalysisUpdateData;
            /**
             *
             * Gets or sets the indentation level of the page.
             *
             * [Api set: OneNoteApi 1.1]
             */
            pageLevel?: number;
            /**
             *
             * Gets or sets the title of the page.
             *
             * [Api set: OneNoteApi 1.1]
             */
            title?: string;
        }
        /** An interface for updating data on the PageCollection object, for use in "pageCollection.set({ ... })". */
        export interface PageCollectionUpdateData {
            items?: OneNote.Interfaces.PageData[];
        }
        /** An interface for updating data on the PageContent object, for use in "pageContent.set({ ... })". */
        export interface PageContentUpdateData {
            /**
            *
            * Gets the Image in the PageContent object. Throws an exception if PageContentType is not Image.
            *
            * [Api set: OneNoteApi 1.1]
            */
            image?: OneNote.Interfaces.ImageUpdateData;
            /**
             *
             * Gets or sets the left (X-axis) position of the PageContent object.
             *
             * [Api set: OneNoteApi 1.1]
             */
            left?: number;
            /**
             *
             * Gets or sets the top (Y-axis) position of the PageContent object.
             *
             * [Api set: OneNoteApi 1.1]
             */
            top?: number;
        }
        /** An interface for updating data on the PageContentCollection object, for use in "pageContentCollection.set({ ... })". */
        export interface PageContentCollectionUpdateData {
            items?: OneNote.Interfaces.PageContentData[];
        }
        /** An interface for updating data on the Paragraph object, for use in "paragraph.set({ ... })". */
        export interface ParagraphUpdateData {
            /**
            *
            * Gets the Image object in the Paragraph. Throws an exception if ParagraphType is not Image.
            *
            * [Api set: OneNoteApi 1.1]
            */
            image?: OneNote.Interfaces.ImageUpdateData;
            /**
            *
            * Gets the Table object in the Paragraph. Throws an exception if ParagraphType is not Table.
            *
            * [Api set: OneNoteApi 1.1]
            */
            table?: OneNote.Interfaces.TableUpdateData;
        }
        /** An interface for updating data on the ParagraphCollection object, for use in "paragraphCollection.set({ ... })". */
        export interface ParagraphCollectionUpdateData {
            items?: OneNote.Interfaces.ParagraphData[];
        }
        /** An interface for updating data on the Image object, for use in "image.set({ ... })". */
        export interface ImageUpdateData {
            /**
             *
             * Gets or sets the description of the Image.
             *
             * [Api set: OneNoteApi 1.1]
             */
            description?: string;
            /**
             *
             * Gets or sets the height of the Image layout.
             *
             * [Api set: OneNoteApi 1.1]
             */
            height?: number;
            /**
             *
             * Gets or sets the hyperlink of the Image.
             *
             * [Api set: OneNoteApi 1.1]
             */
            hyperlink?: string;
            /**
             *
             * Gets or sets the width of the Image layout.
             *
             * [Api set: OneNoteApi 1.1]
             */
            width?: number;
        }
        /** An interface for updating data on the Table object, for use in "table.set({ ... })". */
        export interface TableUpdateData {
            /**
             *
             * Gets or sets whether the borders are visible or not. True if they are visible, false if they are hidden.
             *
             * [Api set: OneNoteApi 1.1]
             */
            borderVisible?: boolean;
        }
        /** An interface for updating data on the TableRowCollection object, for use in "tableRowCollection.set({ ... })". */
        export interface TableRowCollectionUpdateData {
            items?: OneNote.Interfaces.TableRowData[];
        }
        /** An interface for updating data on the TableCell object, for use in "tableCell.set({ ... })". */
        export interface TableCellUpdateData {
            /**
             *
             * Gets and sets the shading color of the cell
             *
             * [Api set: OneNoteApi 1.1]
             */
            shadingColor?: string;
        }
        /** An interface for updating data on the TableCellCollection object, for use in "tableCellCollection.set({ ... })". */
        export interface TableCellCollectionUpdateData {
            items?: OneNote.Interfaces.TableCellData[];
        }
        /** An interface describing the data returned by calling "application.toJSON()". */
        export interface ApplicationData {
            /**
            *
            * Gets the collection of notebooks that are open in the OneNote application instance. In OneNote Online, only one notebook at a time is open in the application instance. Read-only.
            *
            * [Api set: OneNoteApi 1.1]
            */
            notebooks?: OneNote.Interfaces.NotebookData[];
        }
        /** An interface describing the data returned by calling "inkAnalysis.toJSON()". */
        export interface InkAnalysisData {
            /**
            *
            * Gets the parent page object. Read-only.
            *
            * [Api set: OneNoteApi 1.1]
            */
            page?: OneNote.Interfaces.PageData;
            /**
             *
             * Gets the ID of the InkAnalysis object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: string;
        }
        /** An interface describing the data returned by calling "inkAnalysisParagraph.toJSON()". */
        export interface InkAnalysisParagraphData {
            /**
            *
            * Reference to the parent InkAnalysisPage. Read-only.
            *
            * [Api set: OneNoteApi 1.1]
            */
            inkAnalysis?: OneNote.Interfaces.InkAnalysisData;
            /**
            *
            * Gets the ink analysis lines in this ink analysis paragraph. Read-only.
            *
            * [Api set: OneNoteApi 1.1]
            */
            lines?: OneNote.Interfaces.InkAnalysisLineData[];
            /**
             *
             * Gets the ID of the InkAnalysisParagraph object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: string;
        }
        /** An interface describing the data returned by calling "inkAnalysisParagraphCollection.toJSON()". */
        export interface InkAnalysisParagraphCollectionData {
            items?: OneNote.Interfaces.InkAnalysisParagraphData[];
        }
        /** An interface describing the data returned by calling "inkAnalysisLine.toJSON()". */
        export interface InkAnalysisLineData {
            /**
            *
            * Reference to the parent InkAnalysisParagraph. Read-only.
            *
            * [Api set: OneNoteApi 1.1]
            */
            paragraph?: OneNote.Interfaces.InkAnalysisParagraphData;
            /**
            *
            * Gets the ink analysis words in this ink analysis line. Read-only.
            *
            * [Api set: OneNoteApi 1.1]
            */
            words?: OneNote.Interfaces.InkAnalysisWordData[];
            /**
             *
             * Gets the ID of the InkAnalysisLine object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: string;
        }
        /** An interface describing the data returned by calling "inkAnalysisLineCollection.toJSON()". */
        export interface InkAnalysisLineCollectionData {
            items?: OneNote.Interfaces.InkAnalysisLineData[];
        }
        /** An interface describing the data returned by calling "inkAnalysisWord.toJSON()". */
        export interface InkAnalysisWordData {
            /**
            *
            * Reference to the parent InkAnalysisLine. Read-only.
            *
            * [Api set: OneNoteApi 1.1]
            */
            line?: OneNote.Interfaces.InkAnalysisLineData;
            /**
             *
             * Gets the ID of the InkAnalysisWord object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: string;
            /**
             *
             * The id of the recognized language in this inkAnalysisWord. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            languageId?: string;
            /**
             *
             * Weak references to the ink strokes that were recognized as part of this ink analysis word. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            strokePointers?: OneNote.InkStrokePointer[];
            /**
             *
             * The words that were recognized in this ink word, in order of likelihood. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            wordAlternates?: string[];
        }
        /** An interface describing the data returned by calling "inkAnalysisWordCollection.toJSON()". */
        export interface InkAnalysisWordCollectionData {
            items?: OneNote.Interfaces.InkAnalysisWordData[];
        }
        /** An interface describing the data returned by calling "floatingInk.toJSON()". */
        export interface FloatingInkData {
            /**
            *
            * Gets the strokes of the FloatingInk object. Read-only.
            *
            * [Api set: OneNoteApi 1.1]
            */
            inkStrokes?: OneNote.Interfaces.InkStrokeData[];
            /**
             *
             * Gets the ID of the FloatingInk object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: string;
        }
        /** An interface describing the data returned by calling "inkStroke.toJSON()". */
        export interface InkStrokeData {
            /**
            *
            * Gets the ID of the InkStroke object. Read-only.
            *
            * [Api set: OneNoteApi 1.1]
            */
            floatingInk?: OneNote.Interfaces.FloatingInkData;
            /**
             *
             * Gets the ID of the InkStroke object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: string;
        }
        /** An interface describing the data returned by calling "inkStrokeCollection.toJSON()". */
        export interface InkStrokeCollectionData {
            items?: OneNote.Interfaces.InkStrokeData[];
        }
        /** An interface describing the data returned by calling "inkWord.toJSON()". */
        export interface InkWordData {
            /**
             *
             * Gets the ID of the InkWord object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: string;
            /**
             *
             * The id of the recognized language in this ink word. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            languageId?: string;
            /**
             *
             * The words that were recognized in this ink word, in order of likelihood. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            wordAlternates?: string[];
        }
        /** An interface describing the data returned by calling "inkWordCollection.toJSON()". */
        export interface InkWordCollectionData {
            items?: OneNote.Interfaces.InkWordData[];
        }
        /** An interface describing the data returned by calling "notebook.toJSON()". */
        export interface NotebookData {
            /**
            *
            * The section groups in the notebook. Read only
            *
            * [Api set: OneNoteApi 1.1]
            */
            sectionGroups?: OneNote.Interfaces.SectionGroupData[];
            /**
            *
            * The the sections of the notebook. Read only
            *
            * [Api set: OneNoteApi 1.1]
            */
            sections?: OneNote.Interfaces.SectionData[];
            /**
             *
             * The url of the site that this notebook is located. Read only
             *
             * [Api set: OneNoteApi 1.1]
             */
            baseUrl?: string;
            /**
             *
             * The client url of the notebook. Read only
             *
             * [Api set: OneNoteApi 1.1]
             */
            clientUrl?: string;
            /**
             *
             * Gets the ID of the notebook. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: string;
            /**
             *
             * True if the Notebook is not created by the user (i.e. 'Misplaced Sections'). Read only
             *
             * [Api set: OneNoteApi 1.2]
             */
            isVirtual?: boolean;
            /**
             *
             * Gets the name of the notebook. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            name?: string;
        }
        /** An interface describing the data returned by calling "notebookCollection.toJSON()". */
        export interface NotebookCollectionData {
            items?: OneNote.Interfaces.NotebookData[];
        }
        /** An interface describing the data returned by calling "sectionGroup.toJSON()". */
        export interface SectionGroupData {
            /**
            *
            * The collection of section groups in the section group. Read only
            *
            * [Api set: OneNoteApi 1.1]
            */
            sectionGroups?: OneNote.Interfaces.SectionGroupData[];
            /**
            *
            * The collection of sections in the section group. Read only
            *
            * [Api set: OneNoteApi 1.1]
            */
            sections?: OneNote.Interfaces.SectionData[];
            /**
             *
             * The client url of the section group. Read only
             *
             * [Api set: OneNoteApi 1.1]
             */
            clientUrl?: string;
            /**
             *
             * Gets the ID of the section group. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: string;
            /**
             *
             * Gets the name of the section group. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            name?: string;
        }
        /** An interface describing the data returned by calling "sectionGroupCollection.toJSON()". */
        export interface SectionGroupCollectionData {
            items?: OneNote.Interfaces.SectionGroupData[];
        }
        /** An interface describing the data returned by calling "section.toJSON()". */
        export interface SectionData {
            /**
            *
            * The collection of pages in the section. Read only
            *
            * [Api set: OneNoteApi 1.1]
            */
            pages?: OneNote.Interfaces.PageData[];
            /**
             *
             * The client url of the section. Read only
             *
             * [Api set: OneNoteApi 1.1]
             */
            clientUrl?: string;
            /**
             *
             * Gets the ID of the section. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: string;
            /**
             *
             * True if this section is encrypted with a password. Read only
             *
             * [Api set: OneNoteApi 1.2]
             */
            isEncrypted?: boolean;
            /**
             *
             * True if this section is locked. Read only
             *
             * [Api set: OneNoteApi 1.2]
             */
            isLocked?: boolean;
            /**
             *
             * Gets the name of the section. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            name?: string;
            /**
             *
             * The web url of the page. Read only
             *
             * [Api set: OneNoteApi 1.1]
             */
            webUrl?: string;
        }
        /** An interface describing the data returned by calling "sectionCollection.toJSON()". */
        export interface SectionCollectionData {
            items?: OneNote.Interfaces.SectionData[];
        }
        /** An interface describing the data returned by calling "page.toJSON()". */
        export interface PageData {
            /**
            *
            * The collection of PageContent objects on the page. Read only
            *
            * [Api set: OneNoteApi 1.1]
            */
            contents?: OneNote.Interfaces.PageContentData[];
            /**
            *
            * Text interpretation for the ink on the page. Returns null if there is no ink analysis information. Read only.
            *
            * [Api set: OneNoteApi 1.1]
            */
            inkAnalysisOrNull?: OneNote.Interfaces.InkAnalysisData;
            /**
             *
             * Gets the ClassNotebookPageSource to the page.
             *
             * [Api set: OneNoteApi 1.1]
             */
            classNotebookPageSource?: string;
            /**
             *
             * The client url of the page. Read only
             *
             * [Api set: OneNoteApi 1.1]
             */
            clientUrl?: string;
            /**
             *
             * Gets the ID of the page. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: string;
            /**
             *
             * Gets or sets the indentation level of the page.
             *
             * [Api set: OneNoteApi 1.1]
             */
            pageLevel?: number;
            /**
             *
             * Gets or sets the title of the page.
             *
             * [Api set: OneNoteApi 1.1]
             */
            title?: string;
            /**
             *
             * The web url of the page. Read only
             *
             * [Api set: OneNoteApi 1.1]
             */
            webUrl?: string;
        }
        /** An interface describing the data returned by calling "pageCollection.toJSON()". */
        export interface PageCollectionData {
            items?: OneNote.Interfaces.PageData[];
        }
        /** An interface describing the data returned by calling "pageContent.toJSON()". */
        export interface PageContentData {
            /**
            *
            * Gets the Image in the PageContent object. Throws an exception if PageContentType is not Image.
            *
            * [Api set: OneNoteApi 1.1]
            */
            image?: OneNote.Interfaces.ImageData;
            /**
            *
            * Gets the ink in the PageContent object. Throws an exception if PageContentType is not Ink.
            *
            * [Api set: OneNoteApi 1.1]
            */
            ink?: OneNote.Interfaces.FloatingInkData;
            /**
            *
            * Gets the Outline in the PageContent object. Throws an exception if PageContentType is not Outline.
            *
            * [Api set: OneNoteApi 1.1]
            */
            outline?: OneNote.Interfaces.OutlineData;
            /**
             *
             * Gets the ID of the PageContent object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: string;
            /**
             *
             * Gets or sets the left (X-axis) position of the PageContent object.
             *
             * [Api set: OneNoteApi 1.1]
             */
            left?: number;
            /**
             *
             * Gets or sets the top (Y-axis) position of the PageContent object.
             *
             * [Api set: OneNoteApi 1.1]
             */
            top?: number;
            /**
             *
             * Gets the type of the PageContent object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            type?: OneNote.PageContentType | "Outline" | "Image" | "Ink" | "Other";
        }
        /** An interface describing the data returned by calling "pageContentCollection.toJSON()". */
        export interface PageContentCollectionData {
            items?: OneNote.Interfaces.PageContentData[];
        }
        /** An interface describing the data returned by calling "outline.toJSON()". */
        export interface OutlineData {
            /**
            *
            * Gets the collection of Paragraph objects in the Outline. Read-only.
            *
            * [Api set: OneNoteApi 1.1]
            */
            paragraphs?: OneNote.Interfaces.ParagraphData[];
            /**
             *
             * Gets the ID of the Outline object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: string;
        }
        /** An interface describing the data returned by calling "paragraph.toJSON()". */
        export interface ParagraphData {
            /**
            *
            * Gets the Image object in the Paragraph. Throws an exception if ParagraphType is not Image. Read-only.
            *
            * [Api set: OneNoteApi 1.1]
            */
            image?: OneNote.Interfaces.ImageData;
            /**
            *
            * Gets the Ink collection in the Paragraph. Throws an exception if ParagraphType is not Ink. Read-only.
            *
            * [Api set: OneNoteApi 1.1]
            */
            inkWords?: OneNote.Interfaces.InkWordData[];
            /**
            *
            * The collection of paragraphs under this paragraph. Read only
            *
            * [Api set: OneNoteApi 1.1]
            */
            paragraphs?: OneNote.Interfaces.ParagraphData[];
            /**
            *
            * Gets the RichText object in the Paragraph. Throws an exception if ParagraphType is not RichText. Read-only
            *
            * [Api set: OneNoteApi 1.1]
            */
            richText?: OneNote.Interfaces.RichTextData;
            /**
            *
            * Gets the Table object in the Paragraph. Throws an exception if ParagraphType is not Table. Read-only.
            *
            * [Api set: OneNoteApi 1.1]
            */
            table?: OneNote.Interfaces.TableData;
            /**
             *
             * Gets the ID of the Paragraph object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: string;
            /**
             *
             * Gets the type of the Paragraph object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            type?: OneNote.ParagraphType | "RichText" | "Image" | "Table" | "Ink" | "Other";
        }
        /** An interface describing the data returned by calling "paragraphCollection.toJSON()". */
        export interface ParagraphCollectionData {
            items?: OneNote.Interfaces.ParagraphData[];
        }
        /** An interface describing the data returned by calling "noteTag.toJSON()". */
        export interface NoteTagData {
            /**
             *
             * Gets the Id of the NoteTag object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: string;
            /**
             *
             * Gets the status of the NoteTag object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            status?: OneNote.NoteTagStatus | "Unknown" | "Normal" | "Completed" | "Disabled" | "OutlookTask" | "TaskNotSyncedYet" | "TaskRemoved";
            /**
             *
             * Gets the type of the NoteTag object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            type?: OneNote.NoteTagType | "Unknown" | "ToDo" | "Important" | "Question" | "Contact" | "Address" | "PhoneNumber" | "Website" | "Idea" | "Critical" | "ToDoPriority1" | "ToDoPriority2";
        }
        /** An interface describing the data returned by calling "richText.toJSON()". */
        export interface RichTextData {
            /**
             *
             * Gets the ID of the RichText object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: string;
            /**
             *
             * The language id of the text. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            languageId?: string;
            /**
             *
             * Gets the text content of the RichText object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            text?: string;
        }
        /** An interface describing the data returned by calling "image.toJSON()". */
        export interface ImageData {
            /**
             *
             * Gets or sets the description of the Image.
             *
             * [Api set: OneNoteApi 1.1]
             */
            description?: string;
            /**
             *
             * Gets or sets the height of the Image layout.
             *
             * [Api set: OneNoteApi 1.1]
             */
            height?: number;
            /**
             *
             * Gets or sets the hyperlink of the Image.
             *
             * [Api set: OneNoteApi 1.1]
             */
            hyperlink?: string;
            /**
             *
             * Gets the ID of the Image object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: string;
            /**
             *
             * Gets the data obtained by OCR (Optical Character Recognition) of this Image, such as OCR text and language.
             *
             * [Api set: OneNoteApi 1.1]
             */
            ocrData?: OneNote.ImageOcrData;
            /**
             *
             * Gets or sets the width of the Image layout.
             *
             * [Api set: OneNoteApi 1.1]
             */
            width?: number;
        }
        /** An interface describing the data returned by calling "table.toJSON()". */
        export interface TableData {
            /**
            *
            * Gets all of the table rows. Read-only.
            *
            * [Api set: OneNoteApi 1.1]
            */
            rows?: OneNote.Interfaces.TableRowData[];
            /**
             *
             * Gets or sets whether the borders are visible or not. True if they are visible, false if they are hidden.
             *
             * [Api set: OneNoteApi 1.1]
             */
            borderVisible?: boolean;
            /**
             *
             * Gets the number of columns in the table.
             *
             * [Api set: OneNoteApi 1.1]
             */
            columnCount?: number;
            /**
             *
             * Gets the ID of the table. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: string;
            /**
             *
             * Gets the number of rows in the table.
             *
             * [Api set: OneNoteApi 1.1]
             */
            rowCount?: number;
        }
        /** An interface describing the data returned by calling "tableRow.toJSON()". */
        export interface TableRowData {
            /**
            *
            * Gets the cells in the row. Read-only.
            *
            * [Api set: OneNoteApi 1.1]
            */
            cells?: OneNote.Interfaces.TableCellData[];
            /**
             *
             * Gets the number of cells in the row. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            cellCount?: number;
            /**
             *
             * Gets the ID of the row. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: string;
            /**
             *
             * Gets the index of the row in its parent table. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            rowIndex?: number;
        }
        /** An interface describing the data returned by calling "tableRowCollection.toJSON()". */
        export interface TableRowCollectionData {
            items?: OneNote.Interfaces.TableRowData[];
        }
        /** An interface describing the data returned by calling "tableCell.toJSON()". */
        export interface TableCellData {
            /**
            *
            * Gets the collection of Paragraph objects in the TableCell. Read-only.
            *
            * [Api set: OneNoteApi 1.1]
            */
            paragraphs?: OneNote.Interfaces.ParagraphData[];
            /**
             *
             * Gets the index of the cell in its row. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            cellIndex?: number;
            /**
             *
             * Gets the ID of the cell. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: string;
            /**
             *
             * Gets the index of the cell's row in the table. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            rowIndex?: number;
            /**
             *
             * Gets and sets the shading color of the cell
             *
             * [Api set: OneNoteApi 1.1]
             */
            shadingColor?: string;
        }
        /** An interface describing the data returned by calling "tableCellCollection.toJSON()". */
        export interface TableCellCollectionData {
            items?: OneNote.Interfaces.TableCellData[];
        }
        /**
         *
         * Represents the top-level object that contains all globally addressable OneNote objects such as notebooks, the active notebook, and the active section.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface ApplicationLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets the collection of notebooks that are open in the OneNote application instance. In OneNote Online, only one notebook at a time is open in the application instance.
            *
            * [Api set: OneNoteApi 1.1]
            */
            notebooks?: OneNote.Interfaces.NotebookCollectionLoadOptions;
        }
        /**
         *
         * Represents ink analysis data for a given set of ink strokes.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface InkAnalysisLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets the parent page object.
            *
            * [Api set: OneNoteApi 1.1]
            */
            page?: OneNote.Interfaces.PageLoadOptions;
            /**
             *
             * Gets the ID of the InkAnalysis object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
        }
        /**
         *
         * Represents ink analysis data for an identified paragraph formed by ink strokes.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface InkAnalysisParagraphLoadOptions {
            $all?: boolean;
            /**
            *
            * Reference to the parent InkAnalysisPage.
            *
            * [Api set: OneNoteApi 1.1]
            */
            inkAnalysis?: OneNote.Interfaces.InkAnalysisLoadOptions;
            /**
            *
            * Gets the ink analysis lines in this ink analysis paragraph.
            *
            * [Api set: OneNoteApi 1.1]
            */
            lines?: OneNote.Interfaces.InkAnalysisLineCollectionLoadOptions;
            /**
             *
             * Gets the ID of the InkAnalysisParagraph object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
        }
        /**
         *
         * Represents a collection of InkAnalysisParagraph objects.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface InkAnalysisParagraphCollectionLoadOptions {
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Reference to the parent InkAnalysisPage.
            *
            * [Api set: OneNoteApi 1.1]
            */
            inkAnalysis?: OneNote.Interfaces.InkAnalysisLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the ink analysis lines in this ink analysis paragraph.
            *
            * [Api set: OneNoteApi 1.1]
            */
            lines?: OneNote.Interfaces.InkAnalysisLineCollectionLoadOptions;
            /**
             *
             * For EACH ITEM in the collection: Gets the ID of the InkAnalysisParagraph object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
        }
        /**
         *
         * Represents ink analysis data for an identified text line formed by ink strokes.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface InkAnalysisLineLoadOptions {
            $all?: boolean;
            /**
            *
            * Reference to the parent InkAnalysisParagraph.
            *
            * [Api set: OneNoteApi 1.1]
            */
            paragraph?: OneNote.Interfaces.InkAnalysisParagraphLoadOptions;
            /**
            *
            * Gets the ink analysis words in this ink analysis line.
            *
            * [Api set: OneNoteApi 1.1]
            */
            words?: OneNote.Interfaces.InkAnalysisWordCollectionLoadOptions;
            /**
             *
             * Gets the ID of the InkAnalysisLine object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
        }
        /**
         *
         * Represents a collection of InkAnalysisLine objects.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface InkAnalysisLineCollectionLoadOptions {
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Reference to the parent InkAnalysisParagraph.
            *
            * [Api set: OneNoteApi 1.1]
            */
            paragraph?: OneNote.Interfaces.InkAnalysisParagraphLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the ink analysis words in this ink analysis line.
            *
            * [Api set: OneNoteApi 1.1]
            */
            words?: OneNote.Interfaces.InkAnalysisWordCollectionLoadOptions;
            /**
             *
             * For EACH ITEM in the collection: Gets the ID of the InkAnalysisLine object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
        }
        /**
         *
         * Represents ink analysis data for an identified word formed by ink strokes.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface InkAnalysisWordLoadOptions {
            $all?: boolean;
            /**
            *
            * Reference to the parent InkAnalysisLine.
            *
            * [Api set: OneNoteApi 1.1]
            */
            line?: OneNote.Interfaces.InkAnalysisLineLoadOptions;
            /**
             *
             * Gets the ID of the InkAnalysisWord object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * The id of the recognized language in this inkAnalysisWord. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            languageId?: boolean;
            /**
             *
             * Weak references to the ink strokes that were recognized as part of this ink analysis word. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            strokePointers?: boolean;
            /**
             *
             * The words that were recognized in this ink word, in order of likelihood. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            wordAlternates?: boolean;
        }
        /**
         *
         * Represents a collection of InkAnalysisWord objects.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface InkAnalysisWordCollectionLoadOptions {
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Reference to the parent InkAnalysisLine.
            *
            * [Api set: OneNoteApi 1.1]
            */
            line?: OneNote.Interfaces.InkAnalysisLineLoadOptions;
            /**
             *
             * For EACH ITEM in the collection: Gets the ID of the InkAnalysisWord object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * For EACH ITEM in the collection: The id of the recognized language in this inkAnalysisWord. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            languageId?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Weak references to the ink strokes that were recognized as part of this ink analysis word. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            strokePointers?: boolean;
            /**
             *
             * For EACH ITEM in the collection: The words that were recognized in this ink word, in order of likelihood. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            wordAlternates?: boolean;
        }
        /**
         *
         * Represents a group of ink strokes.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface FloatingInkLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets the strokes of the FloatingInk object.
            *
            * [Api set: OneNoteApi 1.1]
            */
            inkStrokes?: OneNote.Interfaces.InkStrokeCollectionLoadOptions;
            /**
            *
            * Gets the PageContent parent of the FloatingInk object.
            *
            * [Api set: OneNoteApi 1.1]
            */
            pageContent?: OneNote.Interfaces.PageContentLoadOptions;
            /**
             *
             * Gets the ID of the FloatingInk object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
        }
        /**
         *
         * Represents a single stroke of ink.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface InkStrokeLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets the ID of the InkStroke object.
            *
            * [Api set: OneNoteApi 1.1]
            */
            floatingInk?: OneNote.Interfaces.FloatingInkLoadOptions;
            /**
             *
             * Gets the ID of the InkStroke object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
        }
        /**
         *
         * Represents a collection of InkStroke objects.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface InkStrokeCollectionLoadOptions {
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Gets the ID of the InkStroke object.
            *
            * [Api set: OneNoteApi 1.1]
            */
            floatingInk?: OneNote.Interfaces.FloatingInkLoadOptions;
            /**
             *
             * For EACH ITEM in the collection: Gets the ID of the InkStroke object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
        }
        /**
         *
         * A container for the ink in a word in a paragraph.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface InkWordLoadOptions {
            $all?: boolean;
            /**
            *
            * The parent paragraph containing the ink word.
            *
            * [Api set: OneNoteApi 1.1]
            */
            paragraph?: OneNote.Interfaces.ParagraphLoadOptions;
            /**
             *
             * Gets the ID of the InkWord object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * The id of the recognized language in this ink word. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            languageId?: boolean;
            /**
             *
             * The words that were recognized in this ink word, in order of likelihood. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            wordAlternates?: boolean;
        }
        /**
         *
         * Represents a collection of InkWord objects.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface InkWordCollectionLoadOptions {
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: The parent paragraph containing the ink word.
            *
            * [Api set: OneNoteApi 1.1]
            */
            paragraph?: OneNote.Interfaces.ParagraphLoadOptions;
            /**
             *
             * For EACH ITEM in the collection: Gets the ID of the InkWord object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * For EACH ITEM in the collection: The id of the recognized language in this ink word. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            languageId?: boolean;
            /**
             *
             * For EACH ITEM in the collection: The words that were recognized in this ink word, in order of likelihood. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            wordAlternates?: boolean;
        }
        /**
         *
         * Represents a OneNote notebook. Notebooks contain section groups and sections.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface NotebookLoadOptions {
            $all?: boolean;
            /**
            *
            * The section groups in the notebook. Read only
            *
            * [Api set: OneNoteApi 1.1]
            */
            sectionGroups?: OneNote.Interfaces.SectionGroupCollectionLoadOptions;
            /**
            *
            * The the sections of the notebook. Read only
            *
            * [Api set: OneNoteApi 1.1]
            */
            sections?: OneNote.Interfaces.SectionCollectionLoadOptions;
            /**
             *
             * The url of the site that this notebook is located. Read only
             *
             * [Api set: OneNoteApi 1.1]
             */
            baseUrl?: boolean;
            /**
             *
             * The client url of the notebook. Read only
             *
             * [Api set: OneNoteApi 1.1]
             */
            clientUrl?: boolean;
            /**
             *
             * Gets the ID of the notebook. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * True if the Notebook is not created by the user (i.e. 'Misplaced Sections'). Read only
             *
             * [Api set: OneNoteApi 1.2]
             */
            isVirtual?: boolean;
            /**
             *
             * Gets the name of the notebook. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            name?: boolean;
        }
        /**
         *
         * Represents a collection of notebooks.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface NotebookCollectionLoadOptions {
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: The section groups in the notebook. Read only
            *
            * [Api set: OneNoteApi 1.1]
            */
            sectionGroups?: OneNote.Interfaces.SectionGroupCollectionLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: The the sections of the notebook. Read only
            *
            * [Api set: OneNoteApi 1.1]
            */
            sections?: OneNote.Interfaces.SectionCollectionLoadOptions;
            /**
             *
             * For EACH ITEM in the collection: The url of the site that this notebook is located. Read only
             *
             * [Api set: OneNoteApi 1.1]
             */
            baseUrl?: boolean;
            /**
             *
             * For EACH ITEM in the collection: The client url of the notebook. Read only
             *
             * [Api set: OneNoteApi 1.1]
             */
            clientUrl?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the ID of the notebook. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * For EACH ITEM in the collection: True if the Notebook is not created by the user (i.e. 'Misplaced Sections'). Read only
             *
             * [Api set: OneNoteApi 1.2]
             */
            isVirtual?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the name of the notebook. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            name?: boolean;
        }
        /**
         *
         * Represents a OneNote section group. Section groups can contain sections and other section groups.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface SectionGroupLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets the notebook that contains the section group.
            *
            * [Api set: OneNoteApi 1.1]
            */
            notebook?: OneNote.Interfaces.NotebookLoadOptions;
            /**
            *
            * Gets the section group that contains the section group. Throws ItemNotFound if the section group is a direct child of the notebook.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentSectionGroup?: OneNote.Interfaces.SectionGroupLoadOptions;
            /**
            *
            * Gets the section group that contains the section group. Returns null if the section group is a direct child of the notebook.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentSectionGroupOrNull?: OneNote.Interfaces.SectionGroupLoadOptions;
            /**
            *
            * The collection of section groups in the section group. Read only
            *
            * [Api set: OneNoteApi 1.1]
            */
            sectionGroups?: OneNote.Interfaces.SectionGroupCollectionLoadOptions;
            /**
            *
            * The collection of sections in the section group. Read only
            *
            * [Api set: OneNoteApi 1.1]
            */
            sections?: OneNote.Interfaces.SectionCollectionLoadOptions;
            /**
             *
             * The client url of the section group. Read only
             *
             * [Api set: OneNoteApi 1.1]
             */
            clientUrl?: boolean;
            /**
             *
             * Gets the ID of the section group. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * Gets the name of the section group. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            name?: boolean;
        }
        /**
         *
         * Represents a collection of section groups.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface SectionGroupCollectionLoadOptions {
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Gets the notebook that contains the section group.
            *
            * [Api set: OneNoteApi 1.1]
            */
            notebook?: OneNote.Interfaces.NotebookLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the section group that contains the section group. Throws ItemNotFound if the section group is a direct child of the notebook.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentSectionGroup?: OneNote.Interfaces.SectionGroupLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the section group that contains the section group. Returns null if the section group is a direct child of the notebook.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentSectionGroupOrNull?: OneNote.Interfaces.SectionGroupLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: The collection of section groups in the section group. Read only
            *
            * [Api set: OneNoteApi 1.1]
            */
            sectionGroups?: OneNote.Interfaces.SectionGroupCollectionLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: The collection of sections in the section group. Read only
            *
            * [Api set: OneNoteApi 1.1]
            */
            sections?: OneNote.Interfaces.SectionCollectionLoadOptions;
            /**
             *
             * For EACH ITEM in the collection: The client url of the section group. Read only
             *
             * [Api set: OneNoteApi 1.1]
             */
            clientUrl?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the ID of the section group. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the name of the section group. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            name?: boolean;
        }
        /**
         *
         * Represents a OneNote section. Sections can contain pages.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface SectionLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets the notebook that contains the section.
            *
            * [Api set: OneNoteApi 1.1]
            */
            notebook?: OneNote.Interfaces.NotebookLoadOptions;
            /**
            *
            * The collection of pages in the section. Read only
            *
            * [Api set: OneNoteApi 1.1]
            */
            pages?: OneNote.Interfaces.PageCollectionLoadOptions;
            /**
            *
            * Gets the section group that contains the section. Throws ItemNotFound if the section is a direct child of the notebook.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentSectionGroup?: OneNote.Interfaces.SectionGroupLoadOptions;
            /**
            *
            * Gets the section group that contains the section. Returns null if the section is a direct child of the notebook.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentSectionGroupOrNull?: OneNote.Interfaces.SectionGroupLoadOptions;
            /**
             *
             * The client url of the section. Read only
             *
             * [Api set: OneNoteApi 1.1]
             */
            clientUrl?: boolean;
            /**
             *
             * Gets the ID of the section. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * True if this section is encrypted with a password. Read only
             *
             * [Api set: OneNoteApi 1.2]
             */
            isEncrypted?: boolean;
            /**
             *
             * True if this section is locked. Read only
             *
             * [Api set: OneNoteApi 1.2]
             */
            isLocked?: boolean;
            /**
             *
             * Gets the name of the section. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            name?: boolean;
            /**
             *
             * The web url of the page. Read only
             *
             * [Api set: OneNoteApi 1.1]
             */
            webUrl?: boolean;
        }
        /**
         *
         * Represents a collection of sections.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface SectionCollectionLoadOptions {
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Gets the notebook that contains the section.
            *
            * [Api set: OneNoteApi 1.1]
            */
            notebook?: OneNote.Interfaces.NotebookLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: The collection of pages in the section. Read only
            *
            * [Api set: OneNoteApi 1.1]
            */
            pages?: OneNote.Interfaces.PageCollectionLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the section group that contains the section. Throws ItemNotFound if the section is a direct child of the notebook.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentSectionGroup?: OneNote.Interfaces.SectionGroupLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the section group that contains the section. Returns null if the section is a direct child of the notebook.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentSectionGroupOrNull?: OneNote.Interfaces.SectionGroupLoadOptions;
            /**
             *
             * For EACH ITEM in the collection: The client url of the section. Read only
             *
             * [Api set: OneNoteApi 1.1]
             */
            clientUrl?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the ID of the section. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * For EACH ITEM in the collection: True if this section is encrypted with a password. Read only
             *
             * [Api set: OneNoteApi 1.2]
             */
            isEncrypted?: boolean;
            /**
             *
             * For EACH ITEM in the collection: True if this section is locked. Read only
             *
             * [Api set: OneNoteApi 1.2]
             */
            isLocked?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the name of the section. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            name?: boolean;
            /**
             *
             * For EACH ITEM in the collection: The web url of the page. Read only
             *
             * [Api set: OneNoteApi 1.1]
             */
            webUrl?: boolean;
        }
        /**
         *
         * Represents a OneNote page.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface PageLoadOptions {
            $all?: boolean;
            /**
            *
            * The collection of PageContent objects on the page. Read only
            *
            * [Api set: OneNoteApi 1.1]
            */
            contents?: OneNote.Interfaces.PageContentCollectionLoadOptions;
            /**
            *
            * Text interpretation for the ink on the page. Returns null if there is no ink analysis information. Read only.
            *
            * [Api set: OneNoteApi 1.1]
            */
            inkAnalysisOrNull?: OneNote.Interfaces.InkAnalysisLoadOptions;
            /**
            *
            * Gets the section that contains the page.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentSection?: OneNote.Interfaces.SectionLoadOptions;
            /**
             *
             * Gets the ClassNotebookPageSource to the page.
             *
             * [Api set: OneNoteApi 1.1]
             */
            classNotebookPageSource?: boolean;
            /**
             *
             * The client url of the page. Read only
             *
             * [Api set: OneNoteApi 1.1]
             */
            clientUrl?: boolean;
            /**
             *
             * Gets the ID of the page. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * Gets or sets the indentation level of the page.
             *
             * [Api set: OneNoteApi 1.1]
             */
            pageLevel?: boolean;
            /**
             *
             * Gets or sets the title of the page.
             *
             * [Api set: OneNoteApi 1.1]
             */
            title?: boolean;
            /**
             *
             * The web url of the page. Read only
             *
             * [Api set: OneNoteApi 1.1]
             */
            webUrl?: boolean;
        }
        /**
         *
         * Represents a collection of pages.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface PageCollectionLoadOptions {
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: The collection of PageContent objects on the page. Read only
            *
            * [Api set: OneNoteApi 1.1]
            */
            contents?: OneNote.Interfaces.PageContentCollectionLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Text interpretation for the ink on the page. Returns null if there is no ink analysis information. Read only.
            *
            * [Api set: OneNoteApi 1.1]
            */
            inkAnalysisOrNull?: OneNote.Interfaces.InkAnalysisLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the section that contains the page.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentSection?: OneNote.Interfaces.SectionLoadOptions;
            /**
             *
             * For EACH ITEM in the collection: Gets the ClassNotebookPageSource to the page.
             *
             * [Api set: OneNoteApi 1.1]
             */
            classNotebookPageSource?: boolean;
            /**
             *
             * For EACH ITEM in the collection: The client url of the page. Read only
             *
             * [Api set: OneNoteApi 1.1]
             */
            clientUrl?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the ID of the page. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the indentation level of the page.
             *
             * [Api set: OneNoteApi 1.1]
             */
            pageLevel?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the title of the page.
             *
             * [Api set: OneNoteApi 1.1]
             */
            title?: boolean;
            /**
             *
             * For EACH ITEM in the collection: The web url of the page. Read only
             *
             * [Api set: OneNoteApi 1.1]
             */
            webUrl?: boolean;
        }
        /**
         *
         * Represents a region on a page that contains top-level content types such as Outline or Image. A PageContent object can be assigned an XY position.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface PageContentLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets the Image in the PageContent object. Throws an exception if PageContentType is not Image.
            *
            * [Api set: OneNoteApi 1.1]
            */
            image?: OneNote.Interfaces.ImageLoadOptions;
            /**
            *
            * Gets the ink in the PageContent object. Throws an exception if PageContentType is not Ink.
            *
            * [Api set: OneNoteApi 1.1]
            */
            ink?: OneNote.Interfaces.FloatingInkLoadOptions;
            /**
            *
            * Gets the Outline in the PageContent object. Throws an exception if PageContentType is not Outline.
            *
            * [Api set: OneNoteApi 1.1]
            */
            outline?: OneNote.Interfaces.OutlineLoadOptions;
            /**
            *
            * Gets the page that contains the PageContent object.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentPage?: OneNote.Interfaces.PageLoadOptions;
            /**
             *
             * Gets the ID of the PageContent object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * Gets or sets the left (X-axis) position of the PageContent object.
             *
             * [Api set: OneNoteApi 1.1]
             */
            left?: boolean;
            /**
             *
             * Gets or sets the top (Y-axis) position of the PageContent object.
             *
             * [Api set: OneNoteApi 1.1]
             */
            top?: boolean;
            /**
             *
             * Gets the type of the PageContent object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            type?: boolean;
        }
        /**
         *
         * Represents the contents of a page, as a collection of PageContent objects.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface PageContentCollectionLoadOptions {
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Gets the Image in the PageContent object. Throws an exception if PageContentType is not Image.
            *
            * [Api set: OneNoteApi 1.1]
            */
            image?: OneNote.Interfaces.ImageLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the ink in the PageContent object. Throws an exception if PageContentType is not Ink.
            *
            * [Api set: OneNoteApi 1.1]
            */
            ink?: OneNote.Interfaces.FloatingInkLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the Outline in the PageContent object. Throws an exception if PageContentType is not Outline.
            *
            * [Api set: OneNoteApi 1.1]
            */
            outline?: OneNote.Interfaces.OutlineLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the page that contains the PageContent object.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentPage?: OneNote.Interfaces.PageLoadOptions;
            /**
             *
             * For EACH ITEM in the collection: Gets the ID of the PageContent object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the left (X-axis) position of the PageContent object.
             *
             * [Api set: OneNoteApi 1.1]
             */
            left?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the top (Y-axis) position of the PageContent object.
             *
             * [Api set: OneNoteApi 1.1]
             */
            top?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the type of the PageContent object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            type?: boolean;
        }
        /**
         *
         * Represents a container for Paragraph objects.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface OutlineLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets the PageContent object that contains the Outline. This object defines the position of the Outline on the page.
            *
            * [Api set: OneNoteApi 1.1]
            */
            pageContent?: OneNote.Interfaces.PageContentLoadOptions;
            /**
            *
            * Gets the collection of Paragraph objects in the Outline.
            *
            * [Api set: OneNoteApi 1.1]
            */
            paragraphs?: OneNote.Interfaces.ParagraphCollectionLoadOptions;
            /**
             *
             * Gets the ID of the Outline object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
        }
        /**
         *
         * A container for the visible content on a page. A Paragraph can contain any one ParagraphType type of content.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface ParagraphLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets the Image object in the Paragraph. Throws an exception if ParagraphType is not Image.
            *
            * [Api set: OneNoteApi 1.1]
            */
            image?: OneNote.Interfaces.ImageLoadOptions;
            /**
            *
            * Gets the Ink collection in the Paragraph. Throws an exception if ParagraphType is not Ink.
            *
            * [Api set: OneNoteApi 1.1]
            */
            inkWords?: OneNote.Interfaces.InkWordCollectionLoadOptions;
            /**
            *
            * Gets the Outline object that contains the Paragraph.
            *
            * [Api set: OneNoteApi 1.1]
            */
            outline?: OneNote.Interfaces.OutlineLoadOptions;
            /**
            *
            * The collection of paragraphs under this paragraph. Read only
            *
            * [Api set: OneNoteApi 1.1]
            */
            paragraphs?: OneNote.Interfaces.ParagraphCollectionLoadOptions;
            /**
            *
            * Gets the parent paragraph object. Throws if a parent paragraph does not exist.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentParagraph?: OneNote.Interfaces.ParagraphLoadOptions;
            /**
            *
            * Gets the parent paragraph object. Returns null if a parent paragraph does not exist.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentParagraphOrNull?: OneNote.Interfaces.ParagraphLoadOptions;
            /**
            *
            * Gets the TableCell object that contains the Paragraph if one exists. If parent is not a TableCell, throws ItemNotFound.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentTableCell?: OneNote.Interfaces.TableCellLoadOptions;
            /**
            *
            * Gets the TableCell object that contains the Paragraph if one exists. If parent is not a TableCell, returns null.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentTableCellOrNull?: OneNote.Interfaces.TableCellLoadOptions;
            /**
            *
            * Gets the RichText object in the Paragraph. Throws an exception if ParagraphType is not RichText.
            *
            * [Api set: OneNoteApi 1.1]
            */
            richText?: OneNote.Interfaces.RichTextLoadOptions;
            /**
            *
            * Gets the Table object in the Paragraph. Throws an exception if ParagraphType is not Table.
            *
            * [Api set: OneNoteApi 1.1]
            */
            table?: OneNote.Interfaces.TableLoadOptions;
            /**
             *
             * Gets the ID of the Paragraph object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * Gets the type of the Paragraph object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            type?: boolean;
        }
        /**
         *
         * Represents a collection of Paragraph objects.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface ParagraphCollectionLoadOptions {
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Gets the Image object in the Paragraph. Throws an exception if ParagraphType is not Image.
            *
            * [Api set: OneNoteApi 1.1]
            */
            image?: OneNote.Interfaces.ImageLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the Ink collection in the Paragraph. Throws an exception if ParagraphType is not Ink.
            *
            * [Api set: OneNoteApi 1.1]
            */
            inkWords?: OneNote.Interfaces.InkWordCollectionLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the Outline object that contains the Paragraph.
            *
            * [Api set: OneNoteApi 1.1]
            */
            outline?: OneNote.Interfaces.OutlineLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: The collection of paragraphs under this paragraph. Read only
            *
            * [Api set: OneNoteApi 1.1]
            */
            paragraphs?: OneNote.Interfaces.ParagraphCollectionLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the parent paragraph object. Throws if a parent paragraph does not exist.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentParagraph?: OneNote.Interfaces.ParagraphLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the parent paragraph object. Returns null if a parent paragraph does not exist.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentParagraphOrNull?: OneNote.Interfaces.ParagraphLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the TableCell object that contains the Paragraph if one exists. If parent is not a TableCell, throws ItemNotFound.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentTableCell?: OneNote.Interfaces.TableCellLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the TableCell object that contains the Paragraph if one exists. If parent is not a TableCell, returns null.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentTableCellOrNull?: OneNote.Interfaces.TableCellLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the RichText object in the Paragraph. Throws an exception if ParagraphType is not RichText.
            *
            * [Api set: OneNoteApi 1.1]
            */
            richText?: OneNote.Interfaces.RichTextLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the Table object in the Paragraph. Throws an exception if ParagraphType is not Table.
            *
            * [Api set: OneNoteApi 1.1]
            */
            table?: OneNote.Interfaces.TableLoadOptions;
            /**
             *
             * For EACH ITEM in the collection: Gets the ID of the Paragraph object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the type of the Paragraph object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            type?: boolean;
        }
        /**
         *
         * A container for the NoteTag in a paragraph.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface NoteTagLoadOptions {
            $all?: boolean;
            /**
             *
             * Gets the Id of the NoteTag object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * Gets the status of the NoteTag object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            status?: boolean;
            /**
             *
             * Gets the type of the NoteTag object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            type?: boolean;
        }
        /**
         *
         * Represents a RichText object in a Paragraph.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface RichTextLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets the Paragraph object that contains the RichText object.
            *
            * [Api set: OneNoteApi 1.1]
            */
            paragraph?: OneNote.Interfaces.ParagraphLoadOptions;
            /**
             *
             * Gets the ID of the RichText object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * The language id of the text. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            languageId?: boolean;
            /**
             *
             * Gets the text content of the RichText object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            text?: boolean;
        }
        /**
         *
         * Represents an Image. An Image can be a direct child of a PageContent object or a Paragraph object.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface ImageLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets the PageContent object that contains the Image. Throws if the Image is not a direct child of a PageContent. This object defines the position of the Image on the page.
            *
            * [Api set: OneNoteApi 1.1]
            */
            pageContent?: OneNote.Interfaces.PageContentLoadOptions;
            /**
            *
            * Gets the Paragraph object that contains the Image. Throws if the Image is not a direct child of a Paragraph.
            *
            * [Api set: OneNoteApi 1.1]
            */
            paragraph?: OneNote.Interfaces.ParagraphLoadOptions;
            /**
             *
             * Gets or sets the description of the Image.
             *
             * [Api set: OneNoteApi 1.1]
             */
            description?: boolean;
            /**
             *
             * Gets or sets the height of the Image layout.
             *
             * [Api set: OneNoteApi 1.1]
             */
            height?: boolean;
            /**
             *
             * Gets or sets the hyperlink of the Image.
             *
             * [Api set: OneNoteApi 1.1]
             */
            hyperlink?: boolean;
            /**
             *
             * Gets the ID of the Image object. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * Gets the data obtained by OCR (Optical Character Recognition) of this Image, such as OCR text and language.
             *
             * [Api set: OneNoteApi 1.1]
             */
            ocrData?: boolean;
            /**
             *
             * Gets or sets the width of the Image layout.
             *
             * [Api set: OneNoteApi 1.1]
             */
            width?: boolean;
        }
        /**
         *
         * Represents a table in a OneNote page.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface TableLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets the Paragraph object that contains the Table object.
            *
            * [Api set: OneNoteApi 1.1]
            */
            paragraph?: OneNote.Interfaces.ParagraphLoadOptions;
            /**
            *
            * Gets all of the table rows.
            *
            * [Api set: OneNoteApi 1.1]
            */
            rows?: OneNote.Interfaces.TableRowCollectionLoadOptions;
            /**
             *
             * Gets or sets whether the borders are visible or not. True if they are visible, false if they are hidden.
             *
             * [Api set: OneNoteApi 1.1]
             */
            borderVisible?: boolean;
            /**
             *
             * Gets the number of columns in the table.
             *
             * [Api set: OneNoteApi 1.1]
             */
            columnCount?: boolean;
            /**
             *
             * Gets the ID of the table. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * Gets the number of rows in the table.
             *
             * [Api set: OneNoteApi 1.1]
             */
            rowCount?: boolean;
        }
        /**
         *
         * Represents a row in a table.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface TableRowLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets the cells in the row.
            *
            * [Api set: OneNoteApi 1.1]
            */
            cells?: OneNote.Interfaces.TableCellCollectionLoadOptions;
            /**
            *
            * Gets the parent table.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentTable?: OneNote.Interfaces.TableLoadOptions;
            /**
             *
             * Gets the number of cells in the row. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            cellCount?: boolean;
            /**
             *
             * Gets the ID of the row. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * Gets the index of the row in its parent table. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            rowIndex?: boolean;
        }
        /**
         *
         * Contains a collection of TableRow objects.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface TableRowCollectionLoadOptions {
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Gets the cells in the row.
            *
            * [Api set: OneNoteApi 1.1]
            */
            cells?: OneNote.Interfaces.TableCellCollectionLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the parent table.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentTable?: OneNote.Interfaces.TableLoadOptions;
            /**
             *
             * For EACH ITEM in the collection: Gets the number of cells in the row. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            cellCount?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the ID of the row. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the index of the row in its parent table. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            rowIndex?: boolean;
        }
        /**
         *
         * Represents a cell in a OneNote table.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface TableCellLoadOptions {
            $all?: boolean;
            /**
            *
            * Gets the collection of Paragraph objects in the TableCell.
            *
            * [Api set: OneNoteApi 1.1]
            */
            paragraphs?: OneNote.Interfaces.ParagraphCollectionLoadOptions;
            /**
            *
            * Gets the parent row of the cell.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentRow?: OneNote.Interfaces.TableRowLoadOptions;
            /**
             *
             * Gets the index of the cell in its row. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            cellIndex?: boolean;
            /**
             *
             * Gets the ID of the cell. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * Gets the index of the cell's row in the table. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            rowIndex?: boolean;
            /**
             *
             * Gets and sets the shading color of the cell
             *
             * [Api set: OneNoteApi 1.1]
             */
            shadingColor?: boolean;
        }
        /**
         *
         * Contains a collection of TableCell objects.
         *
         * [Api set: OneNoteApi 1.1]
         */
        export interface TableCellCollectionLoadOptions {
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Gets the collection of Paragraph objects in the TableCell.
            *
            * [Api set: OneNoteApi 1.1]
            */
            paragraphs?: OneNote.Interfaces.ParagraphCollectionLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the parent row of the cell.
            *
            * [Api set: OneNoteApi 1.1]
            */
            parentRow?: OneNote.Interfaces.TableRowLoadOptions;
            /**
             *
             * For EACH ITEM in the collection: Gets the index of the cell in its row. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            cellIndex?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the ID of the cell. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            id?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the index of the cell's row in the table. Read-only.
             *
             * [Api set: OneNoteApi 1.1]
             */
            rowIndex?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets and sets the shading color of the cell
             *
             * [Api set: OneNoteApi 1.1]
             */
            shadingColor?: boolean;
        }
    }
}
export declare namespace OneNote {
    export class RequestContext extends OfficeExtension.ClientRequestContext {
        constructor(url?: string);
        readonly application: Application;
    }
    /**
     * Executes a batch script that performs actions on the OneNote object model, using a new request context. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param batch - A function that takes in an OneNote.RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the OneNote application. Since the Office add-in and the OneNote application run in two different processes, the request context is required to get access to the OneNote object model from the add-in.
     */
    export function run<T>(batch: (context: OneNote.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the OneNote object model, using the request context of a previously-created API object.
     * @param object - A previously-created API object. The batch will use the same request context as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
     * @param batch - A function that takes in an OneNote.RequestContext and returns a promise (typically, just the result of "context.sync()"). When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     */
    export function run<T>(object: OfficeExtension.ClientObject, batch: (context: OneNote.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the OneNote object model, using the request context of previously-created API objects.
     * @param object - An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared request context, which means that any changes applied to these objects will be picked up by "context.sync()".
     * @param batch - A function that takes in an OneNote.RequestContext and returns a promise (typically, just the result of "context.sync()"). When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     */
    export function run<T>(objects: OfficeExtension.ClientObject[], batch: (context: OneNote.RequestContext) => Promise<T>): Promise<T>;
}


////////////////////////////////////////////////////////////////
/////////////////////// End OneNote APIs ///////////////////////
////////////////////////////////////////////////////////////////