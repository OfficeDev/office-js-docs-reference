import { OfficeExtension } from "../api-extractor-inputs-office/office"
import { Office as Outlook} from "../api-extractor-inputs-outlook/outlook"
////////////////////////////////////////////////////////////////
/////////////////////// Begin Visio APIs ///////////////////////
////////////////////////////////////////////////////////////////

export declare namespace Visio {
    /**
     *
     * Provides information about the document that raised the ShapeAdded event.
     *
     * [Api set:  1.1]
     */
    export interface ShapeAddedEventArgs {
        /**
         *
         * Gets the type of the event. See Visio.EventType for details.
         *
         * [Api set:  1.1]
         */
        type: "ShapeAdded";
        /**
         *
         * ID of the page the shape belongs to.
         *
         * [Api set:  1.1]
         */
        pageId: string;
        /**
         *
         * ID of the shape.
         *
         * [Api set:  1.1]
         */
        shapeId: string;
    }
    /**
     *
     * Provides information about the shape collection that raised the SelectionChanged event.
     *
     * [Api set:  1.1]
     */
    export interface SelectionChangedEventArgs {
        /**
         *
         * Gets the page ID that raised the SelectionChanged event.
         *
         * [Api set:  1.1]
         */
        pageID: number;
        /**
         *
         * Gets the array of shape IDs that raised the SelectionChanged event.
         *
         * [Api set:  1.1]
         */
        shapeIDs: number[];
    }
    /**
     * [Api set:  1.1]
     */
    export class Application extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         *
         * Returns the active document. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly activeDocument: Visio.Document;
        /**
         *
         * Returns the active page. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly activePage: Visio.Page;
        /**
         *
         * Returns the documents collection for a Microsoft Visio instance. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly documents: Visio.DocumentCollection;
        /**
         *
         * Gets or sets if the application is visible.
         *
         * [Api set:  1.1]
         */
        isVisible: boolean;
        /**
         *
         * Specifies the name of the application. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly name: string;
        /**
         *
         * Gets or sets the user name of the application.
         *
         * [Api set:  1.1]
         */
        userName: string;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ApplicationUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Visio.Application): void;
        /**
         *
         * Get the stencil information.
         *
         * [Api set:  1.1]
         *
         * @param stencilName - StencilName represents file name of a stencil.
         * @param includeHiddenMasters - Specifies whether to Include Masters which are Hidden from Visio's UI(like Shapes Panel).The default value is false.
         */
        getStencilInfo(stencilName: string, includeHiddenMasters?: boolean): OfficeExtension.ClientResult<Visio.StencilInfo>;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Visio.Interfaces.ApplicationLoadOptions): Visio.Application;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.Application;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Visio.Application;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Visio.Application object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.ApplicationData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Visio.Interfaces.ApplicationData;
    }
    /**
     * [Api set:  1.1]
     */
    export class Document extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         *
         * Represents a Visio application instance that contains this document. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly application: Visio.Application;
        /**
         *
         * Represents a collection of pages associated with the document. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly pages: Visio.PageCollection;
        /**
         *
         * Gets or sets the description of the document.
         *
         * [Api set:  1.1]
         */
        description: string;
        /**
         *
         * Returns the name of the document, including the drive and path. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly fullName: string;
        /**
         *
         * Returns the ID of the document. Read-only
         *
         * [Api set:  1.1]
         */
        readonly id: string;
        /**
         *
         * Returns the ordinal position of a Document object in the Documents Collection. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly index: number;
        /**
         *
         * Returns the name of the document.
         *
         * [Api set:  1.1]
         */
        readonly name: string;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.DocumentUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Visio.Document): void;
        /**
         *
         * Closes a document.
         *
         * [Api set:  1.1]
         */
        close(): void;
        /**
         *
         * Returns the active page of the document.
         *
         * [Api set:  1.1]
         */
        getActivePage(): Visio.Page;
        /**
         *
         * Set the active page of the document.
         *
         * [Api set:  1.1]
         *
         * @param PageName - Name of the page
         */
        setActivePage(PageName: string): void;
        /**
         *
         * Show or hide a TaskPane.
            This will be consumed by the DV Excel Add-In/Other third-party apps who embed the visio drawing to show/hide the task pane.
         *
         * [Api set:  1.1]
         *
         * @param taskPaneType - Type of the 1st Party TaskPane. It can take values from enum TaskPaneType.
         * @param initialProps - Optional Parameter. This is a generic data structure which would be filled with initial data required to initialize the content of the Taskpane.
         * @param show - Optional Parameter. If it is set to false, it will hide the specified taskpane.
         */
        showTaskPane(taskPaneType: Visio.TaskPaneType, initialProps?: any, show?: boolean): void;
        /**
         *
         * Show or hide a TaskPane.
            This will be consumed by the DV Excel Add-In/Other third-party apps who embed the visio drawing to show/hide the task pane.
         *
         * [Api set:  1.1]
         *
         * @param taskPaneTypeString - Type of the 1st Party TaskPane. It can take values from enum TaskPaneType.
         * @param initialProps - Optional Parameter. This is a generic data structure which would be filled with initial data required to initialize the content of the Taskpane.
         * @param show - Optional Parameter. If it is set to false, it will hide the specified taskpane.
         */
        showTaskPane(taskPaneTypeString: "None" | "DataVisualizerProcessMappings" | "DataVisualizerOrgChartMappings", initialProps?: any, show?: boolean): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Visio.Interfaces.DocumentLoadOptions): Visio.Document;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.Document;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Visio.Document;
        /**
         *
         * Event raised when a data visualizer diagram is created or updated with new mappings and/or data.
         *
         * [Api set:  1.1]
         *
         * @eventproperty
         */
        /**
         *
         * Occurs when the current selection of shapes changes.
         *
         * [Api set:  1.1]
         *
         * @eventproperty
         */
        readonly onSelectionChanged: OfficeExtension.EventHandlers<Visio.SelectionChangedEventArgs>;
        /**
         *
         * Occurs when the shape is added.
         *
         * [Api set:  1.1]
         *
         * @eventproperty
         */
        readonly onShapeAdded: OfficeExtension.EventHandlers<Visio.ShapeAddedEventArgs>;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Visio.Document object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.DocumentData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Visio.Interfaces.DocumentData;
    }
    /**
     * [Api set:  1.1]
     */
    export class DocumentCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Visio.Document[];
        /**
         *
         * Adds a new Document object to the Documents collection.
         *
         * [Api set:  1.1]
         *
         * @param FileName
        - * @returns
         */
        add(FileName: string): Visio.Document;
        /**
         *
         * Returns the number of Documents in Document Collection.
         *
         * [Api set:  1.1]
         * @returns
         */
        getCount(): OfficeExtension.ClientResult<number>;
        getItem(key: number | string): Visio.Document;
        /**
         *
         * Returns an item from a collection.
         *
         * [Api set:  1.1]
         *
         * @param index
        - * @returns
         */
        getItemOrNullObject(index: number): Visio.Document;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Visio.Interfaces.DocumentCollectionLoadOptions & Visio.Interfaces.CollectionLoadOptions): Visio.DocumentCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.DocumentCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Visio.DocumentCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Visio.DocumentCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.DocumentCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Visio.Interfaces.DocumentCollectionData;
    }
    /**
     * [Api set:  1.1]
     */
    export class Page extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         *
         * Returns the instance of Microsoft Visio that is associated with an object. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly application: Visio.Application;
        /**
         *
         * Gets the document object that is assocaited with the page. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly document: Visio.Document;
        /**
         *
         * Shapes at root level, in the page. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly shapes: Visio.ShapeCollection;
        /**
         *
         * Returns the ID of the page. Read-only
         *
         * [Api set:  1.1]
         */
        readonly id: number;
        /**
         *
         * Index of the Page.
         *
         * [Api set:  1.1]
         */
        readonly index: number;
        /**
         *
         * Page name. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly name: string;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.PageUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Visio.Page): void;
        /**
         *
         * Returns a rectangle that tightly encloses the shapes of a page.
         *
         * [Api set:  1.1]
         *
         * @param Flags
        - * @param lpr8Left
        - * @param lpr8Bottom
        - * @param lpr8Right
        - * @param lpr8Top
        - */
        boundingBox(Flags: number, lpr8Left: number, lpr8Bottom: number, lpr8Right: number, lpr8Top: number): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Visio.Interfaces.PageLoadOptions): Visio.Page;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.Page;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Visio.Page;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Visio.Page object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.PageData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Visio.Interfaces.PageData;
    }
    /**
     *
     * Represents a collection of Page objects that are part of the document.
     *
     * [Api set:  1.1]
     */
    export class PageCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Visio.Page[];
        /**
         *
         * Adds a new page to a collection.
         *
         * [Api set:  1.1]
         *
         * @param FileName
        - * @returns
         */
        add(FileName: string): Visio.Page;
        /**
         *
         * Gets the number of pages in the collection.
         *
         * [Api set:  1.1]
         */
        getCount(): OfficeExtension.ClientResult<number>;
        getItem(key: number | string): Visio.Page;
        /**
         *
         * Returns an item from a collection.
         *
         * [Api set:  1.1]
         *
         * @param index
        - * @returns
         */
        getItemOrNullObject(index: number): Visio.Page;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Visio.Interfaces.PageCollectionLoadOptions & Visio.Interfaces.CollectionLoadOptions): Visio.PageCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.PageCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Visio.PageCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Visio.PageCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.PageCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Visio.Interfaces.PageCollectionData;
    }
    /**
     *
     * Represents the Shape class.
     *
     * [Api set:  1.1]
     */
    export class Shape extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         *
         * Returns the instance of Microsoft Visio that is associated with an object. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly application: Visio.Application;
        /**
         *
         * Gets the document object that is assocaited with the shape. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly document: Visio.Document;
        /**
         *
         * Shape's Identifier.
         *
         * [Api set:  1.1]
         */
        readonly id: number;
        /**
         *
         * Indicates whether the shape is a callout shape. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly isCallout: boolean;
        /**
         *
         * Specifes whether a shape is a data graphic callout. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly isDataGraphicCallout: boolean;
        /**
         *
         * Indicates whether a shape is currently open for interactive text editing. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly isOpenForTextEdit: boolean;
        /**
         *
         * Shape's name.
         *
         * [Api set:  1.1]
         */
        readonly name: string;
        /**
         *
         * Returns the type of the object. Read-only.(Shape_Type will give OACR Warning :61721)
         *
         * [Api set:  1.1]
         */
        readonly objType: number;
        /**
         *
         * Shape's Text.
         *
         * [Api set:  1.1]
         */
        readonly text: string;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ShapeUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Visio.Shape): void;
        /**
         *
         * Returns a rectangle that tightly encloses a shape.
         *
         * [Api set:  1.1]
         *
         * @param Flags
        - * @param left
        - * @param bottom
        - * @param right
        - * @param top
        - */
        boundingBox(Flags: number, left: number, bottom: number, right: number, top: number): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Visio.Interfaces.ShapeLoadOptions): Visio.Shape;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.Shape;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Visio.Shape;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Visio.Shape object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.ShapeData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Visio.Interfaces.ShapeData;
    }
    /**
     *
     * Represents the Shape Collection.
     *
     * [Api set:  1.1]
     */
    export class ShapeCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Visio.Shape[];
        /**
         *
         * Adds a new shape to a collection.
         *
         * [Api set:  1.1]
         *
         * @param FileName
        - * @returns
         */
        add(FileName: string): Visio.Shape;
        /**
         *
         * Gets the number of Shapes in the collection.
         *
         * [Api set:  1.1]
         */
        getCount(): OfficeExtension.ClientResult<number>;
        getItem(key: number | string): Visio.Shape;
        /**
         *
         * Returns an item from a collection.
         *
         * [Api set:  1.1]
         *
         * @param index
        - * @returns
         */
        getItemOrNullObject(index: number): Visio.Shape;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Visio.Interfaces.ShapeCollectionLoadOptions & Visio.Interfaces.CollectionLoadOptions): Visio.ShapeCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.ShapeCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Visio.ShapeCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Visio.ShapeCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.ShapeCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Visio.Interfaces.ShapeCollectionData;
    }
    /**
     *
     * Master ionformation.
     *
     * [Api set:  1.1]
     */
    export interface MasterInfo {
        /**
         *
         * Localised Display Name of the Master.
         *
         * [Api set:  1.1]
         */
        name: string;
        /**
         *
         * Master Name.
         *
         * [Api set:  1.1]
         */
        nameU: string;
    }
    /**
     *
     * Stencil Information.
     *
     * [Api set:  1.1]
     */
    export interface StencilInfo {
        /**
         *
         * List of MasterInfo.
         *
         * [Api set:  1.1]
         */
        mastersInfo: Visio.MasterInfo[];
        /**
         *
         * Name represents file name of a stencil.
         *
         * [Api set:  1.1]
         */
        name: string;
        /**
         *
         * Title represents Display Name of the Stencil.
         *
         * [Api set:  1.1]
         */
        title: string;
    }
    /**
     *
     * Message category.
     *
     * [Api set:  1.1]
     */
    enum MessageCategory {
        /**
         *
         * None category.
         *
         */
        none = 0,
        /**
         *
         * Event category.
         *
         */
        event = 65536,
    }
    /**
     *
     * MessageType represents the type of message when event is fired from Host.
     *
     * [Api set:  1.1]
     */
    enum MessageType {
        /**
         *
         * None type.
         *
         */
        none = 0,
        /**
         *
         * Shape Added Event Message.
         *
         */
        shapeAddedEvent = 1,
        /**
         *
         * Selection Changed Event Message.
         *
         */
        selectionChangedEvent = 2,
        /**
         *
         * DataVisualizer diagram operation complete Event Message.
         *
         */
        dataVisualizerDiagramOperationCompletedEvent = 3,
    }
    /**
     *
     * EventType represents the type of the events Host supports.
     *
     * [Api set:  1.1]
     */
    enum EventType {
        /**
         *
         * Shape Added Event.
         *
         */
        shapeAdded = "ShapeAdded",
        /**
         *
         * Selection Changed Event.
         *
         */
        selectionChanged = "SelectionChanged",
        /**
         *
         * DataVisualizer diagram operation complete Event.
         *
         */
        dataVisualizerDiagramOperationCompleted = "DataVisualizerDiagramOperationCompleted",
    }
    /**
     *
     * TaskPaneType represents the types of the First Party TaskPanes that are supported by Host through APIs. Used in case of Show TaskPane API/ TaskPane State Changed Event etc.
     *
     * [Api set:  1.1]
     */
    enum TaskPaneType {
        /**
         *
         * None type.
         *
         */
        none = "None",
        /**
         *
         * Data Visualizer Process Mapping Pane.
         *
         */
        dataVisualizerProcessMappings = "DataVisualizerProcessMappings",
        /**
         *
         * Data Visualizer Organisation Mapping Pane
         *
         */
        dataVisualizerOrgChartMappings = "DataVisualizerOrgChartMappings",
    }
    /**
     *
     * Result of Data Visualizer Diagram operations.
     *
     * [Api set:  1.1]
     */
    enum DataVisualizerDiagramResultType {
        /**
         *
         * Operation is success.
         *
         */
        success = "Success",
        /**
         *
         * Unexpected error during operation.
         *
         */
        unexpected = "Unexpected",
        /**
         *
         * Validation error in operation.
         *
         */
        validationError = "ValidationError",
        /**
         *
         * Conflict error in operation.
         *
         */
        conflictError = "ConflictError",
    }
    /**
     *
     * Type of the Data Visualizer Diagram operation.
     *
     * [Api set:  1.1]
     */
    enum DataVisualizerDiagramOperationType {
        /**
         *
         * unknown operation type.
         *
         */
        unknown = "Unknown",
        /**
         *
         * Creation operation.
         *
         */
        create = "Create",
        /**
         *
         * Update Mappings operation.
         *
         */
        updateMappings = "UpdateMappings",
        /**
         *
         * Update data associated with diagram.
         *
         */
        updateData = "UpdateData",
        /**
         *
         * Update both data and mappings.
         *
         */
        update = "Update",
        /**
         *
         * Delete the diagram content.
         *
         */
        delete = "Delete",
    }
    /**
     *
     * DiagramType for Data Visualizer diagrams.
     *
     * [Api set:  1.1]
     */
    enum DataVisualizerDiagramType {
        /**
         *
         * Unknown.
         *
         */
        unknown = "Unknown",
        /**
         *
         * Basic Flowchart.
         *
         */
        basicFlowchart = "BasicFlowchart",
        /**
         *
         * Horizontal Cross-Functional Flowchart.
         *
         */
        crossFunctionalFlowchart_Horizontal = "CrossFunctionalFlowchart_Horizontal",
        /**
         *
         * Vertical Cross-Functional Flowchart.
         *
         */
        crossFunctionalFlowchart_Vertical = "CrossFunctionalFlowchart_Vertical",
        /**
         *
         * Audit.
         *
         */
        audit = "Audit",
        /**
         *
         * OrgChart.
         *
         */
        orgChart = "OrgChart",
        /**
         *
         * Network.
         *
         */
        network = "Network",
    }
    /**
     *
     * Represents the type of column values.
     *
     * [Api set:  1.1]
     */
    enum ColumnType {
        /**
         *
         * Other.
         *
         */
        unknown = "Unknown",
        /**
         *
         * String values.
         *
         */
        string = "String",
        /**
         *
         * Numerical values.
         *
         */
        number = "Number",
        /**
         *
         * Date.
         *
         */
        date = "Date",
        /**
         *
         * Currency.
         *
         */
        currency = "Currency",
    }
    /**
     *
     * Represents the type of source for the data connection.
     *
     * [Api set:  1.1]
     */
    enum DataSourceType {
        /**
         *
         * Unknown Data Source.
         *
         */
        unknown = "Unknown",
        /**
         *
         * Microsoft Excel workbook.
         *
         */
        excel = "Excel",
    }
    /**
     *
     * Represents the orientation of the Cross Functional Flowchart diagram.
     *
     * [Api set:  1.1]
     */
    enum CrossFunctionalFlowchartOrientation {
        /**
         *
         * Horizontal Cross Functional Flowchart.
         *
         */
        horizontal = "Horizontal",
        /**
         *
         * Vertical Cross Functional Flowchart.
         *
         */
        vertical = "Vertical",
    }
    /**
     *
     * Represents the type of layout.
             Make sure that this enum is same as DVSupportedLayouts visio/Engine/inc/databindingtypes.h
     *
     * [Api set:  1.1]
     */
    enum LayoutVariant {
        /**
         *
         * Invalid layout.
         *
         */
        unknown = "Unknown",
        /**
         *
         * Use the Page default layout.
         *
         */
        pageDefault = "PageDefault",
        /**
         *
         * Use Flowchart with TopToBottom orientation.
         *
         */
        flowchart_TopToBottom = "Flowchart_TopToBottom",
        /**
         *
         * Use Flowchart with LeftToRight orientation.
         *
         */
        flowchart_LeftToRight = "Flowchart_LeftToRight",
        /**
         *
         * Use Radial Layout.
         *
         */
        radial = "Radial",
        /**
         *
         * Use Flowchart with BottomToTop orientation.
         *
         */
        flowchart_BottomToTop = "Flowchart_BottomToTop",
        /**
         *
         * Use Flowchart with RightToLeft orientation.
         *
         */
        flowchart_RightToLeft = "Flowchart_RightToLeft",
        /**
         *
         * Use Circular layout.
         *
         */
        circular = "Circular",
        /**
         *
         * Use WideTree with DownThenRight orientation.
         *
         */
        wideTree_DownThenRight = "WideTree_DownThenRight",
        /**
         *
         * Use WideTree with RightThenDown orientation.
         *
         */
        wideTree_RightThenDown = "WideTree_RightThenDown",
        /**
         *
         * Use WideTree with RightThenUp orientation.
         *
         */
        wideTree_RightThenUp = "WideTree_RightThenUp",
        /**
         *
         * Use WideTree with UpThenRight orientation.
         *
         */
        wideTree_UpThenRight = "WideTree_UpThenRight",
        /**
         *
         * Use WideTree with UpThenLeft orientation.
         *
         */
        wideTree_UpThenLeft = "WideTree_UpThenLeft",
        /**
         *
         * Use WideTree with LeftThenUp orientation.
         *
         */
        wideTree_LeftThenUp = "WideTree_LeftThenUp",
        /**
         *
         * Use WideTree with LeftThenDown orientation.
         *
         */
        wideTree_LeftThenDown = "WideTree_LeftThenDown",
        /**
         *
         * Use WideTree with DownThenLeft orientation.
         *
         */
        wideTree_DownThenLeft = "WideTree_DownThenLeft",
        /**
         *
         * Use ParentDefault layout.
         *
         */
        parentDefault = "ParentDefault",
        /**
         *
         * Use Hierarchy TopToBottomLeft orientation.
         *
         */
        hierarchy_TopToBottomLeft = "Hierarchy_TopToBottomLeft",
        /**
         *
         * Use Hierarchy TopToBottomCenter orientation.
         *
         */
        hierarchy_TopToBottomCenter = "Hierarchy_TopToBottomCenter",
        /**
         *
         * Use Hierarchy TopToBottomRight orientation.
         *
         */
        hierarchy_TopToBottomRight = "Hierarchy_TopToBottomRight",
        /**
         *
         * Use Hierarchy BottomToTopLeft orientation.
         *
         */
        hierarchy_BottomToTopLeft = "Hierarchy_BottomToTopLeft",
        /**
         *
         * Use Hierarchy BottomToTopCenter orientation.
         *
         */
        hierarchy_BottomToTopCenter = "Hierarchy_BottomToTopCenter",
        /**
         *
         * Use Hierarchy BottomToTopRight orientation.
         *
         */
        hierarchy_BottomToTopRight = "Hierarchy_BottomToTopRight",
        /**
         *
         * Use Hierarchy LeftToRightTop orientation.
         *
         */
        hierarchy_LeftToRightTop = "Hierarchy_LeftToRightTop",
        /**
         *
         * Use Hierarchy LeftToRightMiddle orientation.
         *
         */
        hierarchy_LeftToRightMiddle = "Hierarchy_LeftToRightMiddle",
        /**
         *
         * Use Hierarchy LeftToRightBottom orientation.
         *
         */
        hierarchy_LeftToRightBottom = "Hierarchy_LeftToRightBottom",
        /**
         *
         * Use Hierarchy RightToLeftTop orientation.
         *
         */
        hierarchy_RightToLeftTop = "Hierarchy_RightToLeftTop",
        /**
         *
         * Use Hierarchy RightToLeftMiddle orientation.
         *
         */
        hierarchy_RightToLeftMiddle = "Hierarchy_RightToLeftMiddle",
        /**
         *
         * Use Hierarchy RightToLeftBottom orientation.
         *
         */
        hierarchy_RightToLeftBottom = "Hierarchy_RightToLeftBottom",
        /**
         *
         * Use OrgChart with HorizontalCenter orientation.
         *
         */
        orgChart_HorizontalCenter = "OrgChart_HorizontalCenter",
        /**
         *
         * Use OrgChart with HorizontalCenter LeftToRight  orientation.
         *
         */
        orgChart_HorizontalCenter_LeftToRight = "OrgChart_HorizontalCenter_LeftToRight",
        /**
         *
         * Use OrgChart with Hybrid HorizontalCenter and VerticalRight orientation.
         *
         */
        orgChart_Hybrid_HorizontalCenter_VerticalRight = "OrgChart_Hybrid_HorizontalCenter_VerticalRight",
        /**
         *
         * Use OrgChart with SideBySide orientation.
         *
         */
        orgChart_SideBySide = "OrgChart_SideBySide",
    }
    /**
     *
     * Represents the types of data validation error.
     *
     * [Api set:  1.1]
     */
    enum DataValidationErrorType {
        /**
         *
         * No error.
         *
         */
        none = "None",
        /**
         *
         * Data does not have one of the mapped column.
         *
         */
        columnNotMapped = "ColumnNotMapped",
        /**
         *
         * UniqueId column has error.
         *
         */
        uniqueIdColumnError = "UniqueIdColumnError",
        /**
         *
         * Swim-lane column is empty.
         *
         */
        swimlaneColumnError = "SwimlaneColumnError",
        /**
         *
         * Delimiter can not have more then one character.
         *
         */
        delimiterError = "DelimiterError",
        /**
         *
         * Connector column has error.
         *
         */
        connectorColumnError = "ConnectorColumnError",
        /**
         *
         * Connector column is already mapped
            to another setting.
         *
         */
        connectorColumnMappedElsewhere = "ConnectorColumnMappedElsewhere",
        /**
         *
         * Connector label column already mapped
            to other setting.
         *
         */
        connectorLabelColumnMappedElsewhere = "ConnectorLabelColumnMappedElsewhere",
        /**
         *
         * Connector column and connector label column are
            already mapped to other setting.
         *
         */
        connectorColumnAndConnectorLabelMappedElsewhere = "ConnectorColumnAndConnectorLabelMappedElsewhere",
    }
    /**
     *
     * Direction of connector in DataVisualizer diagram.
     *
     * [Api set:  1.1]
     */
    enum ConnectorDirection {
        /**
         *
         * Direction will be from target to source shape.
         *
         */
        fromTarget = "FromTarget",
        /**
         *
         * Direction will be from source to target shape.
         *
         */
        toTarget = "ToTarget",
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
        /** An interface for updating data on the Application object, for use in `application.set({ ... })`. */
        export interface ApplicationUpdateData {
            /**
            *
            * Returns the Active Document.
            *
            * [Api set:  1.1]
            */
            activeDocument?: Visio.Interfaces.DocumentUpdateData;
            /**
            *
            * Returns the Active Page.
            *
            * [Api set:  1.1]
            */
            activePage?: Visio.Interfaces.PageUpdateData;
            /**
             *
             * Gets or sets if the application is visible.
             *
             * [Api set:  1.1]
             */
            isVisible?: boolean;
            /**
             *
             * Gets or sets the user name of the application.
             *
             * [Api set:  1.1]
             */
            userName?: string;
        }
        /** An interface for updating data on the Document object, for use in `document.set({ ... })`. */
        export interface DocumentUpdateData {
            /**
            *
            * Represents a Visio application instance that contains this document.
            *
            * [Api set:  1.1]
            */
            application?: Visio.Interfaces.ApplicationUpdateData;
            /**
             *
             * Gets or sets the description of the document.
             *
             * [Api set:  1.1]
             */
            description?: string;
        }
        /** An interface for updating data on the DocumentCollection object, for use in `documentCollection.set({ ... })`. */
        export interface DocumentCollectionUpdateData {
            items?: Visio.Interfaces.DocumentData[];
        }
        /** An interface for updating data on the Page object, for use in `page.set({ ... })`. */
        export interface PageUpdateData {
            /**
            *
            * Returns the instance of Microsoft Visio that is associated with an object.
            *
            * [Api set:  1.1]
            */
            application?: Visio.Interfaces.ApplicationUpdateData;
            /**
            *
            * Gets the document object that is assocaited with the page.
            *
            * [Api set:  1.1]
            */
            document?: Visio.Interfaces.DocumentUpdateData;
        }
        /** An interface for updating data on the PageCollection object, for use in `pageCollection.set({ ... })`. */
        export interface PageCollectionUpdateData {
            items?: Visio.Interfaces.PageData[];
        }
        /** An interface for updating data on the Shape object, for use in `shape.set({ ... })`. */
        export interface ShapeUpdateData {
            /**
            *
            * Returns the instance of Microsoft Visio that is associated with an object.
            *
            * [Api set:  1.1]
            */
            application?: Visio.Interfaces.ApplicationUpdateData;
            /**
            *
            * Gets the document object that is assocaited with the shape.
            *
            * [Api set:  1.1]
            */
            document?: Visio.Interfaces.DocumentUpdateData;
        }
        /** An interface for updating data on the ShapeCollection object, for use in `shapeCollection.set({ ... })`. */
        export interface ShapeCollectionUpdateData {
            items?: Visio.Interfaces.ShapeData[];
        }
        /** An interface describing the data returned by calling `application.toJSON()`. */
        export interface ApplicationData {
            /**
            *
            * Returns the Active Document. Read-only.
            *
            * [Api set:  1.1]
            */
            activeDocument?: Visio.Interfaces.DocumentData;
            /**
            *
            * Returns the Active Page. Read-only.
            *
            * [Api set:  1.1]
            */
            activePage?: Visio.Interfaces.PageData;
            /**
            *
            * Returns the Documents collection for a Microsoft Visio instance. Read-only.
            *
            * [Api set:  1.1]
            */
            documents?: Visio.Interfaces.DocumentData[];
            /**
             *
             * Gets or sets if the application is visible.
             *
             * [Api set:  1.1]
             */
            isVisible?: boolean;
            /**
             *
             * Specifies the name of the application. Read-only.
             *
             * [Api set:  1.1]
             */
            name?: string;
            /**
             *
             * Gets or sets the user name of the application.
             *
             * [Api set:  1.1]
             */
            userName?: string;
        }
        /** An interface describing the data returned by calling `document.toJSON()`. */
        export interface DocumentData {
            /**
            *
            * Represents a Visio application instance that contains this document. Read-only.
            *
            * [Api set:  1.1]
            */
            application?: Visio.Interfaces.ApplicationData;
            /**
            *
            * Represents a collection of pages associated with the document. Read-only.
            *
            * [Api set:  1.1]
            */
            pages?: Visio.Interfaces.PageData[];
            /**
             *
             * Gets or sets the description of the document.
             *
             * [Api set:  1.1]
             */
            description?: string;
            /**
             *
             * Returns the name of the document, including the drive and path. Read-only.
             *
             * [Api set:  1.1]
             */
            fullName?: string;
            /**
             *
             * Returns the ID of the document. Read-only
             *
             * [Api set:  1.1]
             */
            id?: string;
            /**
             *
             * Returns the ordinal position of a Document object in the Documents Collection. Read-only.
             *
             * [Api set:  1.1]
             */
            index?: number;
            /**
             *
             * Returns the name of the document.
             *
             * [Api set:  1.1]
             */
            name?: string;
        }
        /** An interface describing the data returned by calling `documentCollection.toJSON()`. */
        export interface DocumentCollectionData {
            items?: Visio.Interfaces.DocumentData[];
        }
        /** An interface describing the data returned by calling `page.toJSON()`. */
        export interface PageData {
            /**
            *
            * Returns the instance of Microsoft Visio that is associated with an object. Read-only.
            *
            * [Api set:  1.1]
            */
            application?: Visio.Interfaces.ApplicationData;
            /**
            *
            * Gets the document object that is assocaited with the page. Read-only.
            *
            * [Api set:  1.1]
            */
            document?: Visio.Interfaces.DocumentData;
            /**
            *
            * Shapes at root level, in the page. Read-only.
            *
            * [Api set:  1.1]
            */
            shapes?: Visio.Interfaces.ShapeData[];
            /**
             *
             * Returns the ID of the page. Read-only
             *
             * [Api set:  1.1]
             */
            id?: number;
            /**
             *
             * Index of the Page.
             *
             * [Api set:  1.1]
             */
            index?: number;
            /**
             *
             * Page name. Read-only.
             *
             * [Api set:  1.1]
             */
            name?: string;
        }
        /** An interface describing the data returned by calling `pageCollection.toJSON()`. */
        export interface PageCollectionData {
            items?: Visio.Interfaces.PageData[];
        }
        /** An interface describing the data returned by calling `shape.toJSON()`. */
        export interface ShapeData {
            /**
            *
            * Returns the instance of Microsoft Visio that is associated with an object. Read-only.
            *
            * [Api set:  1.1]
            */
            application?: Visio.Interfaces.ApplicationData;
            /**
            *
            * Gets the document object that is assocaited with the shape. Read-only.
            *
            * [Api set:  1.1]
            */
            document?: Visio.Interfaces.DocumentData;
            /**
             *
             * Shape's Identifier.
             *
             * [Api set:  1.1]
             */
            id?: number;
            /**
             *
             * Indicates whether the shape is a callout shape. Read-only.
             *
             * [Api set:  1.1]
             */
            isCallout?: boolean;
            /**
             *
             * Specifes whether a shape is a data graphic callout. Read-only.
             *
             * [Api set:  1.1]
             */
            isDataGraphicCallout?: boolean;
            /**
             *
             * Indicates whether a shape is currently open for interactive text editing. Read-only.
             *
             * [Api set:  1.1]
             */
            isOpenForTextEdit?: boolean;
            /**
             *
             * Shape's name.
             *
             * [Api set:  1.1]
             */
            name?: string;
            /**
             *
             * Returns the type of the object. Read-only.(Shape_Type will give OACR Warning :61721)
             *
             * [Api set:  1.1]
             */
            objType?: number;
            /**
             *
             * Shape's Text.
             *
             * [Api set:  1.1]
             */
            text?: string;
        }
        /** An interface describing the data returned by calling `shapeCollection.toJSON()`. */
        export interface ShapeCollectionData {
            items?: Visio.Interfaces.ShapeData[];
        }
        /**
         * [Api set:  1.1]
         */
        export interface ApplicationLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            *
            * Returns the Active Document.
            *
            * [Api set:  1.1]
            */
            activeDocument?: Visio.Interfaces.DocumentLoadOptions;
            /**
            *
            * Returns the Active Page.
            *
            * [Api set:  1.1]
            */
            activePage?: Visio.Interfaces.PageLoadOptions;
            /**
             *
             * Gets or sets if the application is visible.
             *
             * [Api set:  1.1]
             */
            isVisible?: boolean;
            /**
             *
             * Specifies the name of the application. Read-only.
             *
             * [Api set:  1.1]
             */
            name?: boolean;
            /**
             *
             * Gets or sets the user name of the application.
             *
             * [Api set:  1.1]
             */
            userName?: boolean;
        }
        /**
         * [Api set:  1.1]
         */
        export interface DocumentLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            *
            * Represents a Visio application instance that contains this document.
            *
            * [Api set:  1.1]
            */
            application?: Visio.Interfaces.ApplicationLoadOptions;
            /**
             *
             * Gets or sets the description of the document.
             *
             * [Api set:  1.1]
             */
            description?: boolean;
            /**
             *
             * Returns the name of the document, including the drive and path. Read-only.
             *
             * [Api set:  1.1]
             */
            fullName?: boolean;
            /**
             *
             * Returns the ID of the document. Read-only
             *
             * [Api set:  1.1]
             */
            id?: boolean;
            /**
             *
             * Returns the ordinal position of a Document object in the Documents Collection. Read-only.
             *
             * [Api set:  1.1]
             */
            index?: boolean;
            /**
             *
             * Returns the name of the document.
             *
             * [Api set:  1.1]
             */
            name?: boolean;
        }
        /**
         * [Api set:  1.1]
         */
        export interface DocumentCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Represents a Visio application instance that contains this document.
            *
            * [Api set:  1.1]
            */
            application?: Visio.Interfaces.ApplicationLoadOptions;
            /**
             *
             * For EACH ITEM in the collection: Gets or sets the description of the document.
             *
             * [Api set:  1.1]
             */
            description?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Returns the name of the document, including the drive and path. Read-only.
             *
             * [Api set:  1.1]
             */
            fullName?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Returns the ID of the document. Read-only
             *
             * [Api set:  1.1]
             */
            id?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Returns the ordinal position of a Document object in the Documents Collection. Read-only.
             *
             * [Api set:  1.1]
             */
            index?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Returns the name of the document.
             *
             * [Api set:  1.1]
             */
            name?: boolean;
        }
        /**
         * [Api set:  1.1]
         */
        export interface PageLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            *
            * Returns the instance of Microsoft Visio that is associated with an object.
            *
            * [Api set:  1.1]
            */
            application?: Visio.Interfaces.ApplicationLoadOptions;
            /**
            *
            * Gets the document object that is assocaited with the page.
            *
            * [Api set:  1.1]
            */
            document?: Visio.Interfaces.DocumentLoadOptions;
            /**
             *
             * Returns the ID of the page. Read-only
             *
             * [Api set:  1.1]
             */
            id?: boolean;
            /**
             *
             * Index of the Page.
             *
             * [Api set:  1.1]
             */
            index?: boolean;
            /**
             *
             * Page name. Read-only.
             *
             * [Api set:  1.1]
             */
            name?: boolean;
        }
        /**
         *
         * Represents a collection of Page objects that are part of the Document.
         *
         * [Api set:  1.1]
         */
        export interface PageCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Returns the instance of Microsoft Visio that is associated with an object.
            *
            * [Api set:  1.1]
            */
            application?: Visio.Interfaces.ApplicationLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the document object that is assocaited with the page.
            *
            * [Api set:  1.1]
            */
            document?: Visio.Interfaces.DocumentLoadOptions;
            /**
             *
             * For EACH ITEM in the collection: Returns the ID of the page. Read-only
             *
             * [Api set:  1.1]
             */
            id?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Index of the Page.
             *
             * [Api set:  1.1]
             */
            index?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Page name. Read-only.
             *
             * [Api set:  1.1]
             */
            name?: boolean;
        }
        /**
         *
         * Represents the Shape class.
         *
         * [Api set:  1.1]
         */
        export interface ShapeLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            *
            * Returns the instance of Microsoft Visio that is associated with an object.
            *
            * [Api set:  1.1]
            */
            application?: Visio.Interfaces.ApplicationLoadOptions;
            /**
            *
            * Gets the document object that is assocaited with the shape.
            *
            * [Api set:  1.1]
            */
            document?: Visio.Interfaces.DocumentLoadOptions;
            /**
             *
             * Shape's Identifier.
             *
             * [Api set:  1.1]
             */
            id?: boolean;
            /**
             *
             * Returns true if the shape is bound to data and is part of Data Visualizer diagram. Read-only.
             *
             * [Api set:  1.1]
             */
            isBoundToData?: boolean;
            /**
             *
             * Indicates whether the shape is a callout shape. Read-only.
             *
             * [Api set:  1.1]
             */
            isCallout?: boolean;
            /**
             *
             * Specifes whether a shape is a data graphic callout. Read-only.
             *
             * [Api set:  1.1]
             */
            isDataGraphicCallout?: boolean;
            /**
             *
             * Indicates whether a shape is currently open for interactive text editing. Read-only.
             *
             * [Api set:  1.1]
             */
            isOpenForTextEdit?: boolean;
            /**
             *
             * Shape's name.
             *
             * [Api set:  1.1]
             */
            name?: boolean;
            /**
             *
             * Returns the type of the object. Read-only.(Shape_Type will give OACR Warning :61721)
             *
             * [Api set:  1.1]
             */
            objType?: boolean;
            /**
             *
             * Shape's Text.
             *
             * [Api set:  1.1]
             */
            text?: boolean;
        }
        /**
         *
         * Represents the Shape Collection.
         *
         * [Api set:  1.1]
         */
        export interface ShapeCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Returns the instance of Microsoft Visio that is associated with an object.
            *
            * [Api set:  1.1]
            */
            application?: Visio.Interfaces.ApplicationLoadOptions;
            /**
            *
            * For EACH ITEM in the collection: Gets the document object that is assocaited with the shape.
            *
            * [Api set:  1.1]
            */
            document?: Visio.Interfaces.DocumentLoadOptions;
            /**
             *
             * For EACH ITEM in the collection: Shape's Identifier.
             *
             * [Api set:  1.1]
             */
            id?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Returns true if the shape is bound to data and is part of Data Visualizer diagram. Read-only.
             *
             * [Api set:  1.1]
             */
            isBoundToData?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Indicates whether the shape is a callout shape. Read-only.
             *
             * [Api set:  1.1]
             */
            isCallout?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Specifes whether a shape is a data graphic callout. Read-only.
             *
             * [Api set:  1.1]
             */
            isDataGraphicCallout?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Indicates whether a shape is currently open for interactive text editing. Read-only.
             *
             * [Api set:  1.1]
             */
            isOpenForTextEdit?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Shape's name.
             *
             * [Api set:  1.1]
             */
            name?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Returns the type of the object. Read-only.(Shape_Type will give OACR Warning :61721)
             *
             * [Api set:  1.1]
             */
            objType?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Shape's Text.
             *
             * [Api set:  1.1]
             */
            text?: boolean;
        }
    }
}
export declare namespace Visio {
    /**
     * The RequestContext object facilitates requests to the Visio application. Since the Office add-in and the Visio application run in two different processes, the request context is required to get access to the Visio object model from the add-in.
     */
    export class RequestContext extends OfficeExtension.ClientRequestContext {
        constructor(url?: string | OfficeExtension.EmbeddedSession);
        readonly document: Document;
    }
    /**
 * Executes a batch script that performs actions on the Visio object model, using a new request context. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
 * @param batch - A function that takes in an Visio.RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Visio application. Since the Office add-in and the Visio application run in two different processes, the request context is required to get access to the Visio object model from the add-in.
 */
    export function run<T>(batch: (context: Visio.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the Visio object model, using the request context of a previously-created API object.
     * @param object - A previously-created API object. The batch will use the same request context as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
     * @param batch - A function that takes in an Visio.RequestContext and returns a promise (typically, just the result of "context.sync()"). When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     */
    export function run<T>(object: OfficeExtension.ClientObject | OfficeExtension.EmbeddedSession, batch: (context: Visio.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the Visio object model, using the request context of previously-created API objects.
     * @param objects - An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared request context, which means that any changes applied to these objects will be picked up by "context.sync()".
     * @param batch - A function that takes in a Visio.RequestContext and returns a promise (typically, just the result of "context.sync()"). When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     */
    export function run<T>(objects: OfficeExtension.ClientObject[], batch: (context: Visio.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the Visio object model, using the RequestContext of a previously-created object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param contextObject - A previously-created Visio.RequestContext. This context will get re-used by the batch function (instead of having a new context created). This means that the batch will be able to pick up changes made to existing API objects, if those objects were derived from this same context.
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Visio application. Since the Office add-in and the Visio application run in two different processes, the RequestContext is required to get access to the Visio object model from the add-in.
     * @remarks
     * In addition to this signature, the method also has the following signatures:
     *
     * `run<T>(batch: (context: Visio.RequestContext) => Promise<T>): Promise<T>;`
     *
     * `run<T>(object: OfficeExtension.ClientObject | OfficeExtension.EmbeddedSession, batch: (context: Visio.RequestContext) => Promise<T>): Promise<T>;`
     *
     * `run<T>(objects: OfficeExtension.ClientObject[], batch: (context: Visio.RequestContext) => Promise<T>): Promise<T>;`
     */
    export function run<T>(contextObject: OfficeExtension.ClientRequestContext, batch: (context: Visio.RequestContext) => Promise<T>): Promise<T>;
}


////////////////////////////////////////////////////////////////
//////////////////////// End Visio APIs ////////////////////////
////////////////////////////////////////////////////////////////