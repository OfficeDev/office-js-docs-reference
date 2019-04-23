import { OfficeExtension } from "../api-extractor-inputs-office/office"
import { Office as Outlook} from "../api-extractor-inputs-outlook/outlook"
////////////////////////////////////////////////////////////////
/////////////////////// Begin Visio APIs ///////////////////////
////////////////////////////////////////////////////////////////

export declare namespace Visio {
    /**
     *
     * Provides information about the shape that raised the ShapeMouseEnter event.
     *
     * [Api set:  1.1]
     */
    export interface ShapeMouseEnterEventArgs {
        /**
         *
         * Gets the name of the page which has the shape object that raised the ShapeMouseEnter event.
         *
         * [Api set:  1.1]
         */
        pageName: string;
        /**
         *
         * Gets the name of the shape object that raised the ShapeMouseEnter event.
         *
         * [Api set:  1.1]
         */
        shapeName: string;
    }
    /**
     *
     * Provides information about the shape that raised the ShapeMouseLeave event.
     *
     * [Api set:  1.1]
     */
    export interface ShapeMouseLeaveEventArgs {
        /**
         *
         * Gets the name of the page which has the shape object that raised the ShapeMouseLeave event.
         *
         * [Api set:  1.1]
         */
        pageName: string;
        /**
         *
         * Gets the name of the shape object that raised the ShapeMouseLeave event.
         *
         * [Api set:  1.1]
         */
        shapeName: string;
    }
    /**
     *
     * Provides information about the page that raised the PageLoadComplete event.
     *
     * [Api set:  1.1]
     */
    export interface PageLoadCompleteEventArgs {
        /**
         *
         * Gets the name of the page that raised the PageLoad event.
         *
         * [Api set:  1.1]
         */
        pageName: string;
        /**
         *
         * Gets the success or failure of the PageLoadComplete event.
         *
         * [Api set:  1.1]
         */
        success: boolean;
    }
    /**
     *
     * Provides information about the document that raised the DataRefreshComplete event.
     *
     * [Api set:  1.1]
     */
    export interface DataRefreshCompleteEventArgs {
        /**
         *
         * Gets the document object that raised the DataRefreshComplete event.
         *
         * [Api set:  1.1]
         */
        document: Visio.Document;
        /**
         *
         * Gets the success or failure of the DataRefreshComplete event.
         *
         * [Api set:  1.1]
         */
        success: boolean;
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
         * Gets the name of the page which has the ShapeCollection object that raised the SelectionChanged event.
         *
         * [Api set:  1.1]
         */
        pageName: string;
        /**
         *
         * Gets the array of shape names that raised the SelectionChanged event.
         *
         * [Api set:  1.1]
         */
        shapeNames: string[];
    }
    /**
     *
     * Provides information about the success or failure of the DocumentLoadComplete event.
     *
     * [Api set:  1.1]
     */
    export interface DocumentLoadCompleteEventArgs {
        /**
         *
         * Gets the success or failure of the DocumentLoadComplete event.
         *
         * [Api set:  1.1]
         */
        success: boolean;
    }
    /**
     *
     * Provides information about the page that raised the PageRenderComplete event.
     *
     * [Api set:  1.1]
     */
    export interface PageRenderCompleteEventArgs {
        /**
         *
         * Gets the name of the page that raised the PageLoad event.
         *
         * [Api set:  1.1]
         */
        pageName: string;
        /**
         *
         * Gets the success/failure of the PageRender event.
         *
         * [Api set:  1.1]
         */
        success: boolean;
    }
    /**
     *
     * Represents the Application.
     *
     * [Api set:  1.1]
     */
    export class Application extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Show or hide the iFrame application borders.
         *
         * [Api set:  1.1]
         */
        showBorders: boolean;
        /**
         *
         * Show or hide the standard toolbars.
         *
         * [Api set:  1.1]
         */
        showToolbars: boolean;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: Visio.Application): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ApplicationUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Visio.Application): void;
        /**
         *
         * Sets the visibility of a specific toolbar in the application.
         *
         * [Api set:  1.1]
         *
         * @param id - The type of the Toolbar
         * @param show - Whether the toolbar is visibile or not.
         */
        showToolbar(id: Visio.ToolBarType, show: boolean): void;
        /**
         *
         * Sets the visibility of a specific toolbar in the application.
         *
         * [Api set:  1.1]
         *
         * @param idString - The type of the Toolbar
         * @param show - Whether the toolbar is visibile or not.
         */
        showToolbar(idString: "CommandBar" | "PageNavigationBar" | "StatusBar", show: boolean): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Visio.Application` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Visio.Application` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Visio.Application` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Visio.Interfaces.ApplicationLoadOptions): Visio.Application;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.Application;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): Visio.Application;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Visio.Application object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.ApplicationData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Visio.Interfaces.ApplicationData;
    }
    /**
     *
     * Represents the Document class.
     *
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
         * Returns the DocumentView object. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly view: Visio.DocumentView;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: Visio.Document): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.DocumentUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Visio.Document): void;
        /**
         *
         * Returns the Active Page of the document.
         *
         * [Api set:  1.1]
         */
        getActivePage(): Visio.Page;
        /**
         *
         * Set the Active Page of the document.
         *
         * [Api set:  1.1]
         *
         * @param PageName - Name of the page
         */
        setActivePage(PageName: string): void;
        /**
         *
         * Triggers the refresh of the data in the Diagram, for all pages.
         *
         * [Api set:  1.1]
         */
        startDataRefresh(): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Visio.Document` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Visio.Document` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Visio.Document` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Visio.Interfaces.DocumentLoadOptions): Visio.Document;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.Document;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): Visio.Document;
        /**
         *
         * Occurs when the data is refreshed in the diagram.
         *
         * [Api set:  1.1]
         *
         * @eventproperty
         */
        readonly onDataRefreshComplete: OfficeExtension.EventHandlers<Visio.DataRefreshCompleteEventArgs>;
        /**
         *
         * Occurs when the Document is loaded, refreshed, or changed.
         *
         * [Api set:  1.1]
         *
         * @eventproperty
         */
        readonly onDocumentLoadComplete: OfficeExtension.EventHandlers<Visio.DocumentLoadCompleteEventArgs>;
        /**
         *
         * Occurs when the page is finished loading.
         *
         * [Api set:  1.1]
         *
         * @eventproperty
         */
        readonly onPageLoadComplete: OfficeExtension.EventHandlers<Visio.PageLoadCompleteEventArgs>;
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
         * Occurs when the user moves the mouse pointer into the bounding box of a shape.
         *
         * [Api set:  1.1]
         *
         * @eventproperty
         */
        readonly onShapeMouseEnter: OfficeExtension.EventHandlers<Visio.ShapeMouseEnterEventArgs>;
        /**
         *
         * Occurs when the user moves the mouse out of the bounding box of a shape.
         *
         * [Api set:  1.1]
         *
         * @eventproperty
         */
        readonly onShapeMouseLeave: OfficeExtension.EventHandlers<Visio.ShapeMouseLeaveEventArgs>;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Visio.Document object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.DocumentData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Visio.Interfaces.DocumentData;
    }
    /**
     *
     * Represents the DocumentView class.
     *
     * [Api set:  1.1]
     */
    export class DocumentView extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Disable Hyperlinks.
         *
         * [Api set:  1.1]
         */
        disableHyperlinks: boolean;
        /**
         *
         * Disable Pan.
         *
         * [Api set:  1.1]
         */
        disablePan: boolean;
        /**
         *
         * Disable PanZoomWindow.
         *
         * [Api set:  1.1]
         */
        disablePanZoomWindow: boolean;
        /**
         *
         * Disable Zoom.
         *
         * [Api set:  1.1]
         */
        disableZoom: boolean;
        /**
         *
         * Hide Diagram Boundary.
         *
         * [Api set:  1.1]
         */
        hideDiagramBoundary: boolean;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: Visio.DocumentView): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.DocumentViewUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Visio.DocumentView): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Visio.DocumentView` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Visio.DocumentView` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Visio.DocumentView` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Visio.Interfaces.DocumentViewLoadOptions): Visio.DocumentView;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.DocumentView;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): Visio.DocumentView;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Visio.DocumentView object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.DocumentViewData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Visio.Interfaces.DocumentViewData;
    }
    /**
     *
     * Represents the Page class.
     *
     * [Api set:  1.1]
     */
    export class Page extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * All shapes in the Page, including subshapes. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly allShapes: Visio.ShapeCollection;
        /**
         *
         * Returns the Comments Collection.  Read-only.
         *
         * [Api set:  1.1]
         */
        readonly comments: Visio.CommentCollection;
        /**
         *
         * All top-level shapes in the Page.Read-only.
         *
         * [Api set:  1.1]
         */
        readonly shapes: Visio.ShapeCollection;
        /**
         *
         * Returns the view of the page. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly view: Visio.PageView;
        /**
         *
         * Returns the height of the page. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly height: number;
        /**
         *
         * Index of the Page. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly index: number;
        /**
         *
         * Whether the page is a background page or not. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly isBackground: boolean;
        /**
         *
         * Page name. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly name: string;
        /**
         *
         * Returns the width of the page. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly width: number;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: Visio.Page): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.PageUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Visio.Page): void;
        /**
         *
         * Set the page as Active Page of the document.
         *
         * [Api set:  1.1]
         */
        activate(): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Visio.Page` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Visio.Page` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Visio.Page` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Visio.Interfaces.PageLoadOptions): Visio.Page;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.Page;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): Visio.Page;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Visio.Page object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.PageData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Visio.Interfaces.PageData;
    }
    /**
     *
     * Represents the PageView class.
     *
     * [Api set:  1.1]
     */
    export class PageView extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Get and set Page's Zoom level. The value can be between 10 and 400 and denotes the percentage of zoom.
         *
         * [Api set:  1.1]
         */
        zoom: number;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: Visio.PageView): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.PageViewUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Visio.PageView): void;
        /**
         *
         * Pans the Visio drawing to place the specified shape in the center of the view.
         *
         * [Api set:  1.1]
         *
         * @param ShapeId - ShapeId to be seen in the center.
         */
        centerViewportOnShape(ShapeId: number): void;
        /**
         *
         * Fit Page to current window.
         *
         * [Api set:  1.1]
         */
        fitToWindow(): void;
        /**
         *
         * Returns the position object that specifies the position of the page in the view.
         *
         * [Api set:  1.1]
         */
        getPosition(): OfficeExtension.ClientResult<Visio.Position>;
        /**
         *
         * Represents the Selection in the page.
         *
         * [Api set:  1.1]
         */
        getSelection(): Visio.Selection;
        /**
         *
         * To check if the shape is in view of the page or not.
         *
         * [Api set:  1.1]
         *
         * @param Shape - Shape to be checked.
         */
        isShapeInViewport(Shape: Visio.Shape): OfficeExtension.ClientResult<boolean>;
        /**
         *
         * Sets the position of the page in the view.
         *
         * [Api set:  1.1]
         *
         * @param Position - Position object that specifies the new position of the page in the view.
         */
        setPosition(Position: Visio.Position): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Visio.PageView` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Visio.PageView` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Visio.PageView` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Visio.Interfaces.PageViewLoadOptions): Visio.PageView;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.PageView;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): Visio.PageView;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Visio.PageView object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.PageViewData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Visio.Interfaces.PageViewData;
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
         * Gets the number of pages in the collection.
         *
         * [Api set:  1.1]
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         *
         * Gets a page using its key (name or Id).
         *
         * [Api set:  1.1]
         *
         * @param key - Key is the name or Id of the page to be retrieved.
         */
        getItem(key: number | string): Visio.Page;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Visio.PageCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Visio.PageCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Visio.PageCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Visio.Interfaces.PageCollectionLoadOptions & Visio.Interfaces.CollectionLoadOptions): Visio.PageCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.PageCollection;
        load(option?: OfficeExtension.LoadOption): Visio.PageCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Visio.PageCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.PageCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Visio.Interfaces.PageCollectionData;
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
         * Gets the number of Shapes in the collection.
         *
         * [Api set:  1.1]
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         *
         * Gets a Shape using its key (name or Index).
         *
         * [Api set:  1.1]
         *
         * @param key - Key is the Name or Index of the shape to be retrieved.
         */
        getItem(key: number | string): Visio.Shape;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Visio.ShapeCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Visio.ShapeCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Visio.ShapeCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Visio.Interfaces.ShapeCollectionLoadOptions & Visio.Interfaces.CollectionLoadOptions): Visio.ShapeCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.ShapeCollection;
        load(option?: OfficeExtension.LoadOption): Visio.ShapeCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Visio.ShapeCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.ShapeCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Visio.Interfaces.ShapeCollectionData;
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
         * Returns the Comments Collection. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly comments: Visio.CommentCollection;
        /**
         *
         * Returns the Hyperlinks collection for a Shape object. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly hyperlinks: Visio.HyperlinkCollection;
        /**
         *
         * Returns the Shape's Data Section. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly shapeDataItems: Visio.ShapeDataItemCollection;
        /**
         *
         * Gets SubShape Collection. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly subShapes: Visio.ShapeCollection;
        /**
         *
         * Returns the view of the shape. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly view: Visio.ShapeView;
        /**
         *
         * Shape's identifier. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly id: number;
        /**
         *
         * Shape's name. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly name: string;
        /**
         *
         * Returns true, if shape is selected. User can set true to select the shape explicitly.
         *
         * [Api set:  1.1]
         */
        select: boolean;
        /**
         *
         * Shape's text. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly text: string;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: Visio.Shape): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ShapeUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Visio.Shape): void;
        /**
         *
         * Returns the BoundingBox object that specifies bounding box of the shape.
         *
         * [Api set:  1.1]
         */
        getBounds(): OfficeExtension.ClientResult<Visio.BoundingBox>;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Visio.Shape` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Visio.Shape` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Visio.Shape` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Visio.Interfaces.ShapeLoadOptions): Visio.Shape;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.Shape;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): Visio.Shape;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Visio.Shape object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.ShapeData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Visio.Interfaces.ShapeData;
    }
    /**
     *
     * Represents the ShapeView class.
     *
     * [Api set:  1.1]
     */
    export class ShapeView extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Represents the highlight around the shape.
         *
         * [Api set:  1.1]
         */
        highlight: Visio.Highlight;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: Visio.ShapeView): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ShapeViewUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Visio.ShapeView): void;
        /**
         *
         * Adds an overlay on top of the shape.
         *
         * [Api set:  1.1]
         *
         * @param OverlayType - An Overlay Type. Can be 'Text' or 'Image'.
         * @param Content - Content of Overlay.
         * @param OverlayHorizontalAlignment - Horizontal Alignment of Overlay. Can be 'Left', 'Center', or 'Right'.
         * @param OverlayVerticalAlignment - Vertical Alignment of Overlay. Can be 'Top', 'Middle', 'Bottom'.
         * @param Width - Overlay Width.
         * @param Height - Overlay Height.
         */
        addOverlay(OverlayType: Visio.OverlayType, Content: string, OverlayHorizontalAlignment: Visio.OverlayHorizontalAlignment, OverlayVerticalAlignment: Visio.OverlayVerticalAlignment, Width: number, Height: number): OfficeExtension.ClientResult<number>;
        /**
         *
         * Adds an overlay on top of the shape.
         *
         * [Api set:  1.1]
         *
         * @param OverlayTypeString - An Overlay Type. Can be 'Text' or 'Image'.
         * @param Content - Content of Overlay.
         * @param OverlayHorizontalAlignment - Horizontal Alignment of Overlay. Can be 'Left', 'Center', or 'Right'.
         * @param OverlayVerticalAlignment - Vertical Alignment of Overlay. Can be 'Top', 'Middle', 'Bottom'.
         * @param Width - Overlay Width.
         * @param Height - Overlay Height.
         */
        addOverlay(OverlayTypeString: "Text" | "Image" | "Html", Content: string, OverlayHorizontalAlignment: "Left" | "Center" | "Right", OverlayVerticalAlignment: "Top" | "Middle" | "Bottom", Width: number, Height: number): OfficeExtension.ClientResult<number>;
        /**
         *
         * Removes particular overlay or all overlays on the Shape.
         *
         * [Api set:  1.1]
         *
         * @param OverlayId - An Overlay Id. Removes the specific overlay id from the shape.
         */
        removeOverlay(OverlayId: number): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Visio.ShapeView` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Visio.ShapeView` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Visio.ShapeView` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Visio.Interfaces.ShapeViewLoadOptions): Visio.ShapeView;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.ShapeView;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): Visio.ShapeView;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Visio.ShapeView object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.ShapeViewData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Visio.Interfaces.ShapeViewData;
    }
    /**
     *
     * Represents the Position of the object in the view.
     *
     * [Api set:  1.1]
     */
    export interface Position {
        /**
         *
         * An integer that specifies the x-coordinate of the object, which is the signed value of the distance in pixels from the viewport's center to the left boundary of the page.
         *
         * [Api set:  1.1]
         */
        x: number;
        /**
         *
         * An integer that specifies the y-coordinate of the object, which is the signed value of the distance in pixels from the viewport's center to the top boundary of the page.
         *
         * [Api set:  1.1]
         */
        y: number;
    }
    /**
     *
     * Represents the BoundingBox of the shape.
     *
     * [Api set:  1.1]
     */
    export interface BoundingBox {
        /**
         *
         * The distance between the top and bottom edges of the bounding box of the shape, excluding any data graphics associated with the shape.
         *
         * [Api set:  1.1]
         */
        height: number;
        /**
         *
         * The distance between the left and right edges of the bounding box of the shape, excluding any data graphics associated with the shape.
         *
         * [Api set:  1.1]
         */
        width: number;
        /**
         *
         * An integer that specifies the x-coordinate of the bounding box.
         *
         * [Api set:  1.1]
         */
        x: number;
        /**
         *
         * An integer that specifies the y-coordinate of the bounding box.
         *
         * [Api set:  1.1]
         */
        y: number;
    }
    /**
     *
     * Represents the highlight data added to the shape.
     *
     * [Api set:  1.1]
     */
    export interface Highlight {
        /**
         *
         * A string that specifies the color of the highlight. It must have the form "#RRGGBB", where each letter represents a hexadecimal digit between 0 and F, and where RR is the red value between 0 and 0xFF (255), GG the green value between 0 and 0xFF (255), and BB is the blue value between 0 and 0xFF (255).
         *
         * [Api set:  1.1]
         */
        color: string;
        /**
         *
         * A positive integer that specifies the width of the highlight's stroke in pixels.
         *
         * [Api set:  1.1]
         */
        width: number;
    }
    /**
     *
     * Represents the ShapeDataItemCollection for a given Shape.
     *
     * [Api set:  1.1]
     */
    export class ShapeDataItemCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: Visio.ShapeDataItem[];
        /**
         *
         * Gets the number of Shape Data Items.
         *
         * [Api set:  1.1]
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         *
         * Gets the ShapeDataItem using its name.
         *
         * [Api set:  1.1]
         *
         * @param key - Key is the name of the ShapeDataItem to be retrieved.
         */
        getItem(key: string): Visio.ShapeDataItem;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Visio.ShapeDataItemCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Visio.ShapeDataItemCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Visio.ShapeDataItemCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Visio.Interfaces.ShapeDataItemCollectionLoadOptions & Visio.Interfaces.CollectionLoadOptions): Visio.ShapeDataItemCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.ShapeDataItemCollection;
        load(option?: OfficeExtension.LoadOption): Visio.ShapeDataItemCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Visio.ShapeDataItemCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.ShapeDataItemCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Visio.Interfaces.ShapeDataItemCollectionData;
    }
    /**
     *
     * Represents the ShapeDataItem.
     *
     * [Api set:  1.1]
     */
    export class ShapeDataItem extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * A string that specifies the format of the shape data item. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly format: string;
        /**
         *
         * A string that specifies the formatted value of the shape data item. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly formattedValue: string;
        /**
         *
         * A string that specifies the label of the shape data item. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly label: string;
        /**
         *
         * A string that specifies the value of the shape data item. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly value: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Visio.ShapeDataItem` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Visio.ShapeDataItem` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Visio.ShapeDataItem` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Visio.Interfaces.ShapeDataItemLoadOptions): Visio.ShapeDataItem;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.ShapeDataItem;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): Visio.ShapeDataItem;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Visio.ShapeDataItem object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.ShapeDataItemData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Visio.Interfaces.ShapeDataItemData;
    }
    /**
     *
     * Represents the Hyperlink Collection.
     *
     * [Api set:  1.1]
     */
    export class HyperlinkCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: Visio.Hyperlink[];
        /**
         *
         * Gets the number of hyperlinks.
         *
         * [Api set:  1.1]
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         *
         * Gets a Hyperlink using its key (name or Id).
         *
         * [Api set:  1.1]
         *
         * @param Key - Key is the name or index of the Hyperlink to be retrieved.
         */
        getItem(Key: number | string): Visio.Hyperlink;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Visio.HyperlinkCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Visio.HyperlinkCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Visio.HyperlinkCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Visio.Interfaces.HyperlinkCollectionLoadOptions & Visio.Interfaces.CollectionLoadOptions): Visio.HyperlinkCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.HyperlinkCollection;
        load(option?: OfficeExtension.LoadOption): Visio.HyperlinkCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Visio.HyperlinkCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.HyperlinkCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Visio.Interfaces.HyperlinkCollectionData;
    }
    /**
     *
     * Represents the Hyperlink.
     *
     * [Api set:  1.1]
     */
    export class Hyperlink extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the address of the Hyperlink object. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly address: string;
        /**
         *
         * Gets the description of a hyperlink. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly description: string;
        /**
         *
         * Gets the extra URL request information used to resolve the hyperlink's URL. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly extraInfo: string;
        /**
         *
         * Gets the sub-address of the Hyperlink object. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly subAddress: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Visio.Hyperlink` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Visio.Hyperlink` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Visio.Hyperlink` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Visio.Interfaces.HyperlinkLoadOptions): Visio.Hyperlink;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.Hyperlink;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): Visio.Hyperlink;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Visio.Hyperlink object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.HyperlinkData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Visio.Interfaces.HyperlinkData;
    }
    /**
     *
     * Represents the CommentCollection for a given Shape.
     *
     * [Api set:  1.1]
     */
    export class CommentCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /** Gets the loaded child items in this collection. */
        readonly items: Visio.Comment[];
        /**
         *
         * Gets the number of Comments.
         *
         * [Api set:  1.1]
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         *
         * Gets the Comment using its name.
         *
         * [Api set:  1.1]
         *
         * @param key - Key is the name of the Comment to be retrieved.
         */
        getItem(key: string): Visio.Comment;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Visio.CommentCollection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Visio.CommentCollection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Visio.CommentCollection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Visio.Interfaces.CommentCollectionLoadOptions & Visio.Interfaces.CollectionLoadOptions): Visio.CommentCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.CommentCollection;
        load(option?: OfficeExtension.LoadOption): Visio.CommentCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Visio.CommentCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.CommentCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Visio.Interfaces.CommentCollectionData;
    }
    /**
     *
     * Represents the Comment.
     *
     * [Api set:  1.1]
     */
    export class Comment extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * A string that specifies the name of the author of the comment.
         *
         * [Api set:  1.1]
         */
        author: string;
        /**
         *
         * A string that specifies the date when the comment was created.
         *
         * [Api set:  1.1]
         */
        date: string;
        /**
         *
         * A string that contains the comment text.
         *
         * [Api set:  1.1]
         */
        text: string;
        /** Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         *
         * @remarks
         *
         * This method has the following additional signature:
         *
         * `set(properties: Visio.Comment): void`
         *
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.CommentUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Visio.Comment): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Visio.Comment` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Visio.Comment` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Visio.Comment` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(option?: Visio.Interfaces.CommentLoadOptions): Visio.Comment;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.Comment;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): Visio.Comment;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Visio.Comment object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.CommentData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Visio.Interfaces.CommentData;
    }
    /**
     *
     * Represents the Selection in the page.
     *
     * [Api set:  1.1]
     */
    export class Selection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext; 
        /**
         *
         * Gets the Shapes of the Selection. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly shapes: Visio.ShapeCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         *
         * @remarks
         *
         * In addition to this signature, this method has the following signatures:
         *
         * `load(option?: string | string[]): Visio.Selection` - Where option is a comma-delimited string or an array of strings that specify the properties to load.
         *
         * `load(option?: { select?: string; expand?: string; }): Visio.Selection` - Where option.select is a comma-delimited string that specifies the properties to load, and options.expand is a comma-delimited string that specifies the navigation properties to load.
         *
         * `load(option?: { select?: string; expand?: string; top?: number; skip?: number }): Visio.Selection` - Only available on collection types. It is similar to the preceding signature. Option.top specifies the maximum number of collection items that can be included in the result. Option.skip specifies the number of items that are to be skipped and not included in the result. If option.top is specified, the result set will start after skipping the specified number of items.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Visio.Selection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         * @param propertyNamesAndPaths - Where propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: { select?: string; expand?: string; }): Visio.Selection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Visio.Selection object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Visio.Interfaces.SelectionData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Visio.Interfaces.SelectionData;
    }
    /**
     *
     * Represents the Horizontal Alignment of the Overlay relative to the shape.
     *
     * [Api set:  1.1]
     */
    enum OverlayHorizontalAlignment {
        /**
         *
         * left
         *
         */
        left = "Left",
        /**
         *
         * center
         *
         */
        center = "Center",
        /**
         *
         * right
         *
         */
        right = "Right",
    }
    /**
     *
     * Represents the Vertical Alignment of the Overlay relative to the shape.
     *
     * [Api set:  1.1]
     */
    enum OverlayVerticalAlignment {
        /**
         *
         * top
         *
         */
        top = "Top",
        /**
         *
         * middle
         *
         */
        middle = "Middle",
        /**
         *
         * bottom
         *
         */
        bottom = "Bottom",
    }
    /**
     *
     * Represents the type of the overlay.
     *
     * [Api set:  1.1]
     */
    enum OverlayType {
        /**
         *
         * text
         *
         */
        text = "Text",
        /**
         *
         * image
         *
         */
        image = "Image",
        /**
         *
         * html
         *
         */
        html = "Html",
    }
    /**
     *
     * Toolbar IDs of the app
     *
     * [Api set:  1.1]
     */
    enum ToolBarType {
        /**
         *
         * CommandBar
         *
         */
        commandBar = "CommandBar",
        /**
         *
         * PageNavigationBar
         *
         */
        pageNavigationBar = "PageNavigationBar",
        /**
         *
         * StatusBar
         *
         */
        statusBar = "StatusBar",
    }
    enum ErrorCodes {
        accessDenied = "AccessDenied",
        generalException = "GeneralException",
        invalidArgument = "InvalidArgument",
        itemNotFound = "ItemNotFound",
        notImplemented = "NotImplemented",
        unsupportedOperation = "UnsupportedOperation",
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
        /** An interface for updating data on the Application object, for use in "application.set({ ... })". */
        export interface ApplicationUpdateData {
            /**
             *
             * Show or hide the iFrame application borders.
             *
             * [Api set:  1.1]
             */
            showBorders?: boolean;
            /**
             *
             * Show or hide the standard toolbars.
             *
             * [Api set:  1.1]
             */
            showToolbars?: boolean;
        }
        /** An interface for updating data on the Document object, for use in "document.set({ ... })". */
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
            * Returns the DocumentView object.
            *
            * [Api set:  1.1]
            */
            view?: Visio.Interfaces.DocumentViewUpdateData;
        }
        /** An interface for updating data on the DocumentView object, for use in "documentView.set({ ... })". */
        export interface DocumentViewUpdateData {
            /**
             *
             * Disable Hyperlinks.
             *
             * [Api set:  1.1]
             */
            disableHyperlinks?: boolean;
            /**
             *
             * Disable Pan.
             *
             * [Api set:  1.1]
             */
            disablePan?: boolean;
            /**
             *
             * Disable PanZoomWindow.
             *
             * [Api set:  1.1]
             */
            disablePanZoomWindow?: boolean;
            /**
             *
             * Disable Zoom.
             *
             * [Api set:  1.1]
             */
            disableZoom?: boolean;
            /**
             *
             * Hide Diagram Boundary.
             *
             * [Api set:  1.1]
             */
            hideDiagramBoundary?: boolean;
        }
        /** An interface for updating data on the Page object, for use in "page.set({ ... })". */
        export interface PageUpdateData {
            /**
            *
            * Returns the view of the page.
            *
            * [Api set:  1.1]
            */
            view?: Visio.Interfaces.PageViewUpdateData;
        }
        /** An interface for updating data on the PageView object, for use in "pageView.set({ ... })". */
        export interface PageViewUpdateData {
            /**
             *
             * Get and set Page's Zoom level. The value can be between 10 and 400 and denotes the percentage of zoom.
             *
             * [Api set:  1.1]
             */
            zoom?: number;
        }
        /** An interface for updating data on the PageCollection object, for use in "pageCollection.set({ ... })". */
        export interface PageCollectionUpdateData {
            items?: Visio.Interfaces.PageData[];
        }
        /** An interface for updating data on the ShapeCollection object, for use in "shapeCollection.set({ ... })". */
        export interface ShapeCollectionUpdateData {
            items?: Visio.Interfaces.ShapeData[];
        }
        /** An interface for updating data on the Shape object, for use in "shape.set({ ... })". */
        export interface ShapeUpdateData {
            /**
            *
            * Returns the view of the shape.
            *
            * [Api set:  1.1]
            */
            view?: Visio.Interfaces.ShapeViewUpdateData;
            /**
             *
             * Returns true, if shape is selected. User can set true to select the shape explicitly.
             *
             * [Api set:  1.1]
             */
            select?: boolean;
        }
        /** An interface for updating data on the ShapeView object, for use in "shapeView.set({ ... })". */
        export interface ShapeViewUpdateData {
            /**
             *
             * Represents the highlight around the shape.
             *
             * [Api set:  1.1]
             */
            highlight?: Visio.Highlight;
        }
        /** An interface for updating data on the ShapeDataItemCollection object, for use in "shapeDataItemCollection.set({ ... })". */
        export interface ShapeDataItemCollectionUpdateData {
            items?: Visio.Interfaces.ShapeDataItemData[];
        }
        /** An interface for updating data on the HyperlinkCollection object, for use in "hyperlinkCollection.set({ ... })". */
        export interface HyperlinkCollectionUpdateData {
            items?: Visio.Interfaces.HyperlinkData[];
        }
        /** An interface for updating data on the CommentCollection object, for use in "commentCollection.set({ ... })". */
        export interface CommentCollectionUpdateData {
            items?: Visio.Interfaces.CommentData[];
        }
        /** An interface for updating data on the Comment object, for use in "comment.set({ ... })". */
        export interface CommentUpdateData {
            /**
             *
             * A string that specifies the name of the author of the comment.
             *
             * [Api set:  1.1]
             */
            author?: string;
            /**
             *
             * A string that specifies the date when the comment was created.
             *
             * [Api set:  1.1]
             */
            date?: string;
            /**
             *
             * A string that contains the comment text.
             *
             * [Api set:  1.1]
             */
            text?: string;
        }
        /** An interface describing the data returned by calling "application.toJSON()". */
        export interface ApplicationData {
            /**
             *
             * Show or hide the iFrame application borders.
             *
             * [Api set:  1.1]
             */
            showBorders?: boolean;
            /**
             *
             * Show or hide the standard toolbars.
             *
             * [Api set:  1.1]
             */
            showToolbars?: boolean;
        }
        /** An interface describing the data returned by calling "document.toJSON()". */
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
            * Returns the DocumentView object. Read-only.
            *
            * [Api set:  1.1]
            */
            view?: Visio.Interfaces.DocumentViewData;
        }
        /** An interface describing the data returned by calling "documentView.toJSON()". */
        export interface DocumentViewData {
            /**
             *
             * Disable Hyperlinks.
             *
             * [Api set:  1.1]
             */
            disableHyperlinks?: boolean;
            /**
             *
             * Disable Pan.
             *
             * [Api set:  1.1]
             */
            disablePan?: boolean;
            /**
             *
             * Disable PanZoomWindow.
             *
             * [Api set:  1.1]
             */
            disablePanZoomWindow?: boolean;
            /**
             *
             * Disable Zoom.
             *
             * [Api set:  1.1]
             */
            disableZoom?: boolean;
            /**
             *
             * Hide Diagram Boundary.
             *
             * [Api set:  1.1]
             */
            hideDiagramBoundary?: boolean;
        }
        /** An interface describing the data returned by calling "page.toJSON()". */
        export interface PageData {
            /**
            *
            * All shapes in the Page, including subshapes. Read-only.
            *
            * [Api set:  1.1]
            */
            allShapes?: Visio.Interfaces.ShapeData[];
            /**
            *
            * Returns the Comments Collection.  Read-only.
            *
            * [Api set:  1.1]
            */
            comments?: Visio.Interfaces.CommentData[];
            /**
            *
            * All top-level shapes in the Page.Read-only.
            *
            * [Api set:  1.1]
            */
            shapes?: Visio.Interfaces.ShapeData[];
            /**
            *
            * Returns the view of the page. Read-only.
            *
            * [Api set:  1.1]
            */
            view?: Visio.Interfaces.PageViewData;
            /**
             *
             * Returns the height of the page. Read-only.
             *
             * [Api set:  1.1]
             */
            height?: number;
            /**
             *
             * Index of the Page. Read-only.
             *
             * [Api set:  1.1]
             */
            index?: number;
            /**
             *
             * Whether the page is a background page or not. Read-only.
             *
             * [Api set:  1.1]
             */
            isBackground?: boolean;
            /**
             *
             * Page name. Read-only.
             *
             * [Api set:  1.1]
             */
            name?: string;
            /**
             *
             * Returns the width of the page. Read-only.
             *
             * [Api set:  1.1]
             */
            width?: number;
        }
        /** An interface describing the data returned by calling "pageView.toJSON()". */
        export interface PageViewData {
            /**
             *
             * Get and set Page's Zoom level. The value can be between 10 and 400 and denotes the percentage of zoom.
             *
             * [Api set:  1.1]
             */
            zoom?: number;
        }
        /** An interface describing the data returned by calling "pageCollection.toJSON()". */
        export interface PageCollectionData {
            items?: Visio.Interfaces.PageData[];
        }
        /** An interface describing the data returned by calling "shapeCollection.toJSON()". */
        export interface ShapeCollectionData {
            items?: Visio.Interfaces.ShapeData[];
        }
        /** An interface describing the data returned by calling "shape.toJSON()". */
        export interface ShapeData {
            /**
            *
            * Returns the Comments Collection. Read-only.
            *
            * [Api set:  1.1]
            */
            comments?: Visio.Interfaces.CommentData[];
            /**
            *
            * Returns the Hyperlinks collection for a Shape object. Read-only.
            *
            * [Api set:  1.1]
            */
            hyperlinks?: Visio.Interfaces.HyperlinkData[];
            /**
            *
            * Returns the Shape's Data Section. Read-only.
            *
            * [Api set:  1.1]
            */
            shapeDataItems?: Visio.Interfaces.ShapeDataItemData[];
            /**
            *
            * Gets SubShape Collection. Read-only.
            *
            * [Api set:  1.1]
            */
            subShapes?: Visio.Interfaces.ShapeData[];
            /**
            *
            * Returns the view of the shape. Read-only.
            *
            * [Api set:  1.1]
            */
            view?: Visio.Interfaces.ShapeViewData;
            /**
             *
             * Shape's identifier. Read-only.
             *
             * [Api set:  1.1]
             */
            id?: number;
            /**
             *
             * Shape's name. Read-only.
             *
             * [Api set:  1.1]
             */
            name?: string;
            /**
             *
             * Returns true, if shape is selected. User can set true to select the shape explicitly.
             *
             * [Api set:  1.1]
             */
            select?: boolean;
            /**
             *
             * Shape's text. Read-only.
             *
             * [Api set:  1.1]
             */
            text?: string;
        }
        /** An interface describing the data returned by calling "shapeView.toJSON()". */
        export interface ShapeViewData {
            /**
             *
             * Represents the highlight around the shape.
             *
             * [Api set:  1.1]
             */
            highlight?: Visio.Highlight;
        }
        /** An interface describing the data returned by calling "shapeDataItemCollection.toJSON()". */
        export interface ShapeDataItemCollectionData {
            items?: Visio.Interfaces.ShapeDataItemData[];
        }
        /** An interface describing the data returned by calling "shapeDataItem.toJSON()". */
        export interface ShapeDataItemData {
            /**
             *
             * A string that specifies the format of the shape data item. Read-only.
             *
             * [Api set:  1.1]
             */
            format?: string;
            /**
             *
             * A string that specifies the formatted value of the shape data item. Read-only.
             *
             * [Api set:  1.1]
             */
            formattedValue?: string;
            /**
             *
             * A string that specifies the label of the shape data item. Read-only.
             *
             * [Api set:  1.1]
             */
            label?: string;
            /**
             *
             * A string that specifies the value of the shape data item. Read-only.
             *
             * [Api set:  1.1]
             */
            value?: string;
        }
        /** An interface describing the data returned by calling "hyperlinkCollection.toJSON()". */
        export interface HyperlinkCollectionData {
            items?: Visio.Interfaces.HyperlinkData[];
        }
        /** An interface describing the data returned by calling "hyperlink.toJSON()". */
        export interface HyperlinkData {
            /**
             *
             * Gets the address of the Hyperlink object. Read-only.
             *
             * [Api set:  1.1]
             */
            address?: string;
            /**
             *
             * Gets the description of a hyperlink. Read-only.
             *
             * [Api set:  1.1]
             */
            description?: string;
            /**
             *
             * Gets the extra URL request information used to resolve the hyperlink's URL. Read-only.
             *
             * [Api set:  1.1]
             */
            extraInfo?: string;
            /**
             *
             * Gets the sub-address of the Hyperlink object. Read-only.
             *
             * [Api set:  1.1]
             */
            subAddress?: string;
        }
        /** An interface describing the data returned by calling "commentCollection.toJSON()". */
        export interface CommentCollectionData {
            items?: Visio.Interfaces.CommentData[];
        }
        /** An interface describing the data returned by calling "comment.toJSON()". */
        export interface CommentData {
            /**
             *
             * A string that specifies the name of the author of the comment.
             *
             * [Api set:  1.1]
             */
            author?: string;
            /**
             *
             * A string that specifies the date when the comment was created.
             *
             * [Api set:  1.1]
             */
            date?: string;
            /**
             *
             * A string that contains the comment text.
             *
             * [Api set:  1.1]
             */
            text?: string;
        }
        /** An interface describing the data returned by calling "selection.toJSON()". */
        export interface SelectionData {
            /**
            *
            * Gets the Shapes of the Selection. Read-only.
            *
            * [Api set:  1.1]
            */
            shapes?: Visio.Interfaces.ShapeData[];
        }
        /**
         *
         * Represents the Application.
         *
         * [Api set:  1.1]
         */
        export interface ApplicationLoadOptions {
            $all?: boolean;
            /**
             *
             * Show or hide the iFrame application borders.
             *
             * [Api set:  1.1]
             */
            showBorders?: boolean;
            /**
             *
             * Show or hide the standard toolbars.
             *
             * [Api set:  1.1]
             */
            showToolbars?: boolean;
        }
        /**
         *
         * Represents the Document class.
         *
         * [Api set:  1.1]
         */
        export interface DocumentLoadOptions {
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
            * Returns the DocumentView object.
            *
            * [Api set:  1.1]
            */
            view?: Visio.Interfaces.DocumentViewLoadOptions;
        }
        /**
         *
         * Represents the DocumentView class.
         *
         * [Api set:  1.1]
         */
        export interface DocumentViewLoadOptions {
            $all?: boolean;
            /**
             *
             * Disable Hyperlinks.
             *
             * [Api set:  1.1]
             */
            disableHyperlinks?: boolean;
            /**
             *
             * Disable Pan.
             *
             * [Api set:  1.1]
             */
            disablePan?: boolean;
            /**
             *
             * Disable PanZoomWindow.
             *
             * [Api set:  1.1]
             */
            disablePanZoomWindow?: boolean;
            /**
             *
             * Disable Zoom.
             *
             * [Api set:  1.1]
             */
            disableZoom?: boolean;
            /**
             *
             * Hide Diagram Boundary.
             *
             * [Api set:  1.1]
             */
            hideDiagramBoundary?: boolean;
        }
        /**
         *
         * Represents the Page class.
         *
         * [Api set:  1.1]
         */
        export interface PageLoadOptions {
            $all?: boolean;
            /**
            *
            * Returns the view of the page.
            *
            * [Api set:  1.1]
            */
            view?: Visio.Interfaces.PageViewLoadOptions;
            /**
             *
             * Returns the height of the page. Read-only.
             *
             * [Api set:  1.1]
             */
            height?: boolean;
            /**
             *
             * Index of the Page. Read-only.
             *
             * [Api set:  1.1]
             */
            index?: boolean;
            /**
             *
             * Whether the page is a background page or not. Read-only.
             *
             * [Api set:  1.1]
             */
            isBackground?: boolean;
            /**
             *
             * Page name. Read-only.
             *
             * [Api set:  1.1]
             */
            name?: boolean;
            /**
             *
             * Returns the width of the page. Read-only.
             *
             * [Api set:  1.1]
             */
            width?: boolean;
        }
        /**
         *
         * Represents the PageView class.
         *
         * [Api set:  1.1]
         */
        export interface PageViewLoadOptions {
            $all?: boolean;
            /**
             *
             * Get and set Page's Zoom level. The value can be between 10 and 400 and denotes the percentage of zoom.
             *
             * [Api set:  1.1]
             */
            zoom?: boolean;
        }
        /**
         *
         * Represents a collection of Page objects that are part of the document.
         *
         * [Api set:  1.1]
         */
        export interface PageCollectionLoadOptions {
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Returns the view of the page.
            *
            * [Api set:  1.1]
            */
            view?: Visio.Interfaces.PageViewLoadOptions;
            /**
             *
             * For EACH ITEM in the collection: Returns the height of the page. Read-only.
             *
             * [Api set:  1.1]
             */
            height?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Index of the Page. Read-only.
             *
             * [Api set:  1.1]
             */
            index?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Whether the page is a background page or not. Read-only.
             *
             * [Api set:  1.1]
             */
            isBackground?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Page name. Read-only.
             *
             * [Api set:  1.1]
             */
            name?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Returns the width of the page. Read-only.
             *
             * [Api set:  1.1]
             */
            width?: boolean;
        }
        /**
         *
         * Represents the Shape Collection.
         *
         * [Api set:  1.1]
         */
        export interface ShapeCollectionLoadOptions {
            $all?: boolean;
            /**
            *
            * For EACH ITEM in the collection: Returns the view of the shape.
            *
            * [Api set:  1.1]
            */
            view?: Visio.Interfaces.ShapeViewLoadOptions;
            /**
             *
             * For EACH ITEM in the collection: Shape's identifier. Read-only.
             *
             * [Api set:  1.1]
             */
            id?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Shape's name. Read-only.
             *
             * [Api set:  1.1]
             */
            name?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Returns true, if shape is selected. User can set true to select the shape explicitly.
             *
             * [Api set:  1.1]
             */
            select?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Shape's text. Read-only.
             *
             * [Api set:  1.1]
             */
            text?: boolean;
        }
        /**
         *
         * Represents the Shape class.
         *
         * [Api set:  1.1]
         */
        export interface ShapeLoadOptions {
            $all?: boolean;
            /**
            *
            * Returns the view of the shape.
            *
            * [Api set:  1.1]
            */
            view?: Visio.Interfaces.ShapeViewLoadOptions;
            /**
             *
             * Shape's identifier. Read-only.
             *
             * [Api set:  1.1]
             */
            id?: boolean;
            /**
             *
             * Shape's name. Read-only.
             *
             * [Api set:  1.1]
             */
            name?: boolean;
            /**
             *
             * Returns true, if shape is selected. User can set true to select the shape explicitly.
             *
             * [Api set:  1.1]
             */
            select?: boolean;
            /**
             *
             * Shape's text. Read-only.
             *
             * [Api set:  1.1]
             */
            text?: boolean;
        }
        /**
         *
         * Represents the ShapeView class.
         *
         * [Api set:  1.1]
         */
        export interface ShapeViewLoadOptions {
            $all?: boolean;
            /**
             *
             * Represents the highlight around the shape.
             *
             * [Api set:  1.1]
             */
            highlight?: boolean;
        }
        /**
         *
         * Represents the ShapeDataItemCollection for a given Shape.
         *
         * [Api set:  1.1]
         */
        export interface ShapeDataItemCollectionLoadOptions {
            $all?: boolean;
            /**
             *
             * For EACH ITEM in the collection: A string that specifies the format of the shape data item. Read-only.
             *
             * [Api set:  1.1]
             */
            format?: boolean;
            /**
             *
             * For EACH ITEM in the collection: A string that specifies the formatted value of the shape data item. Read-only.
             *
             * [Api set:  1.1]
             */
            formattedValue?: boolean;
            /**
             *
             * For EACH ITEM in the collection: A string that specifies the label of the shape data item. Read-only.
             *
             * [Api set:  1.1]
             */
            label?: boolean;
            /**
             *
             * For EACH ITEM in the collection: A string that specifies the value of the shape data item. Read-only.
             *
             * [Api set:  1.1]
             */
            value?: boolean;
        }
        /**
         *
         * Represents the ShapeDataItem.
         *
         * [Api set:  1.1]
         */
        export interface ShapeDataItemLoadOptions {
            $all?: boolean;
            /**
             *
             * A string that specifies the format of the shape data item. Read-only.
             *
             * [Api set:  1.1]
             */
            format?: boolean;
            /**
             *
             * A string that specifies the formatted value of the shape data item. Read-only.
             *
             * [Api set:  1.1]
             */
            formattedValue?: boolean;
            /**
             *
             * A string that specifies the label of the shape data item. Read-only.
             *
             * [Api set:  1.1]
             */
            label?: boolean;
            /**
             *
             * A string that specifies the value of the shape data item. Read-only.
             *
             * [Api set:  1.1]
             */
            value?: boolean;
        }
        /**
         *
         * Represents the Hyperlink Collection.
         *
         * [Api set:  1.1]
         */
        export interface HyperlinkCollectionLoadOptions {
            $all?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the address of the Hyperlink object. Read-only.
             *
             * [Api set:  1.1]
             */
            address?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the description of a hyperlink. Read-only.
             *
             * [Api set:  1.1]
             */
            description?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the extra URL request information used to resolve the hyperlink's URL. Read-only.
             *
             * [Api set:  1.1]
             */
            extraInfo?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the sub-address of the Hyperlink object. Read-only.
             *
             * [Api set:  1.1]
             */
            subAddress?: boolean;
        }
        /**
         *
         * Represents the Hyperlink.
         *
         * [Api set:  1.1]
         */
        export interface HyperlinkLoadOptions {
            $all?: boolean;
            /**
             *
             * Gets the address of the Hyperlink object. Read-only.
             *
             * [Api set:  1.1]
             */
            address?: boolean;
            /**
             *
             * Gets the description of a hyperlink. Read-only.
             *
             * [Api set:  1.1]
             */
            description?: boolean;
            /**
             *
             * Gets the extra URL request information used to resolve the hyperlink's URL. Read-only.
             *
             * [Api set:  1.1]
             */
            extraInfo?: boolean;
            /**
             *
             * Gets the sub-address of the Hyperlink object. Read-only.
             *
             * [Api set:  1.1]
             */
            subAddress?: boolean;
        }
        /**
         *
         * Represents the CommentCollection for a given Shape.
         *
         * [Api set:  1.1]
         */
        export interface CommentCollectionLoadOptions {
            $all?: boolean;
            /**
             *
             * For EACH ITEM in the collection: A string that specifies the name of the author of the comment.
             *
             * [Api set:  1.1]
             */
            author?: boolean;
            /**
             *
             * For EACH ITEM in the collection: A string that specifies the date when the comment was created.
             *
             * [Api set:  1.1]
             */
            date?: boolean;
            /**
             *
             * For EACH ITEM in the collection: A string that contains the comment text.
             *
             * [Api set:  1.1]
             */
            text?: boolean;
        }
        /**
         *
         * Represents the Comment.
         *
         * [Api set:  1.1]
         */
        export interface CommentLoadOptions {
            $all?: boolean;
            /**
             *
             * A string that specifies the name of the author of the comment.
             *
             * [Api set:  1.1]
             */
            author?: boolean;
            /**
             *
             * A string that specifies the date when the comment was created.
             *
             * [Api set:  1.1]
             */
            date?: boolean;
            /**
             *
             * A string that contains the comment text.
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