// Type definitions for Office.js
// Project: http://dev.office.com
// Definitions by: OfficeDev <https://github.com/OfficeDev>, Lance Austin <https://github.com/LanceEA>, Michael Zlatkovsky <https://github.com/Zlatkovsky>, Kim Brandl <https://github.com/kbrandl>, Ricky Kirkham <https://github.com/Rick-Kirkham>
// Definitions: https://github.com/DefinitelyTyped/DefinitelyTyped

/*
office-js
Copyright (c) Microsoft Corporation
*/


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
         * Gets the shape object that raised the ShapeMouseEnter event.
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
         * Gets the shape object that raised the ShapeMouseLeave event.
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
         * Gets the success/failure of the PageLoadComplete event.
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
         * Gets the success/failure of the DataRefreshComplete event.
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
         * Gets the ShapeCollection object that raised the SelectionChanged event.
         *
         * [Api set:  1.1]
         */
        shapeNames: Array<string>;
    }
    /**
     *
     * Provides information about the drawing that raised the DiagramLoadComplete event.
     *
     * [Api set:  1.1]
     */
    export interface DocumentLoadCompleteEventArgs {
        /**
         *
         * Gets the success/failure of the DocumentLoadComplete event.
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
        /**
         *
         * Show/Hide the application borders.
         *
         * [Api set:  1.1]
         */
        showBorders: boolean;
        /**
         *
         * Show or Hide the standard toolbars.
         *
         * [Api set:  1.1]
         */
        showToolbars: boolean;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.ApplicationUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Application): void;
        /**
         *
         * Show or Hide a particular toolbar.
         *
         * [Api set:  1.1]
         */
        showToolbar(id: string, show: boolean): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: Visio.Interfaces.ApplicationLoadOptions): Visio.Application;
        load(option?: string | string[]): Visio.Application;
        load(option?: {
            select?: string;
            expand?: string;
        }): Visio.Application;
        toJSON(): Visio.Interfaces.ApplicationData;
    }
    /**
     *
     * Represents the Document class.
     *
     * [Api set:  1.1]
     */
    export class Document extends OfficeExtension.ClientObject {
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
         * Returns the DocumentView object.
         *
         * [Api set:  1.1]
         */
        readonly view: Visio.DocumentView;
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
         */
        load(option?: Visio.Interfaces.DocumentLoadOptions): Visio.Document;
        load(option?: string | string[]): Visio.Document;
        load(option?: {
            select?: string;
            expand?: string;
        }): Visio.Document;
        /**
         *
         * Occurs when the data is refreshed in the diagram.
         *
         * [Api set:  1.1]
         */
        readonly onDataRefreshComplete: OfficeExtension.EventHandlers<Visio.DataRefreshCompleteEventArgs>;
        /**
         *
         * Occurs when the Document is loaded, refreshed, or changed.
         *
         * [Api set:  1.1]
         */
        readonly onDocumentLoadComplete: OfficeExtension.EventHandlers<Visio.DocumentLoadCompleteEventArgs>;
        /**
         *
         * Occurs when the page is finished loading.
         *
         * [Api set:  1.1]
         */
        readonly onPageLoadComplete: OfficeExtension.EventHandlers<Visio.PageLoadCompleteEventArgs>;
        /**
         *
         * Occurs when the current selection of shapes changes.
         *
         * [Api set:  1.1]
         */
        readonly onSelectionChanged: OfficeExtension.EventHandlers<Visio.SelectionChangedEventArgs>;
        /**
         *
         * Occurs when the user moves the mouse pointer into the bounding box of a shape.
         *
         * [Api set:  1.1]
         */
        readonly onShapeMouseEnter: OfficeExtension.EventHandlers<Visio.ShapeMouseEnterEventArgs>;
        /**
         *
         * Occurs when the user moves the mouse out of the bounding box of a shape.
         *
         * [Api set:  1.1]
         */
        readonly onShapeMouseLeave: OfficeExtension.EventHandlers<Visio.ShapeMouseLeaveEventArgs>;
        toJSON(): Visio.Interfaces.DocumentData;
    }
    /**
     *
     * Represents the DocumentView class.
     *
     * [Api set:  1.1]
     */
    export class DocumentView extends OfficeExtension.ClientObject {
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
         * Disable Zoom.
         *
         * [Api set:  1.1]
         */
        disableZoom: boolean;
        /**
         *
         * Disable Hyperlinks.
         *
         * [Api set:  1.1]
         */
        hideDiagramBoundary: boolean;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.DocumentViewUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: DocumentView): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: Visio.Interfaces.DocumentViewLoadOptions): Visio.DocumentView;
        load(option?: string | string[]): Visio.DocumentView;
        load(option?: {
            select?: string;
            expand?: string;
        }): Visio.DocumentView;
        toJSON(): Visio.Interfaces.DocumentViewData;
    }
    /**
     *
     * Represents the Page class.
     *
     * [Api set:  1.1]
     */
    export class Page extends OfficeExtension.ClientObject {
        /**
         *
         * All shapes in the page. Read-only.
         *
         * [Api set:  1.1]
         */
        readonly allShapes: Visio.ShapeCollection;
        /**
         *
         * Returns the Comments Collection
         *
         * [Api set:  1.1]
         */
        readonly comments: Visio.CommentCollection;
        /**
         *
         * Shapes at root level, in the page. Read-only.
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
         * Index of the Page.
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
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.PageUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Page): void;
        /**
         *
         * Set the page as Active Page of the document.
         *
         * [Api set:  1.1]
         */
        activate(): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: Visio.Interfaces.PageLoadOptions): Visio.Page;
        load(option?: string | string[]): Visio.Page;
        load(option?: {
            select?: string;
            expand?: string;
        }): Visio.Page;
        toJSON(): Visio.Interfaces.PageData;
    }
    /**
     *
     * Represents the PageView class.
     *
     * [Api set:  1.1]
     */
    export class PageView extends OfficeExtension.ClientObject {
        /**
         *
         * Get/Set Page's Zoom level. The value can be between 10 and 400 and denotes the percentage of zoom.
         *
         * [Api set:  1.1]
         */
        zoom: number;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.PageViewUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: PageView): void;
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
         */
        load(option?: Visio.Interfaces.PageViewLoadOptions): Visio.PageView;
        load(option?: string | string[]): Visio.PageView;
        load(option?: {
            select?: string;
            expand?: string;
        }): Visio.PageView;
        toJSON(): Visio.Interfaces.PageViewData;
    }
    /**
     *
     * Represents a collection of Page objects that are part of the document.
     *
     * [Api set:  1.1]
     */
    export class PageCollection extends OfficeExtension.ClientObject {
        /** Gets the loaded child items in this collection. */
        readonly items: Array<Visio.Page>;
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
         */
        load(option?: Visio.Interfaces.PageCollectionLoadOptions & Visio.Interfaces.CollectionLoadOptions): Visio.PageCollection;
        load(option?: string | string[]): Visio.PageCollection;
        load(option?: OfficeExtension.LoadOption): Visio.PageCollection;
        toJSON(): Visio.Interfaces.PageCollectionData;
    }
    /**
     *
     * Represents the Shape Collection.
     *
     * [Api set:  1.1]
     */
    export class ShapeCollection extends OfficeExtension.ClientObject {
        /** Gets the loaded child items in this collection. */
        readonly items: Array<Visio.Shape>;
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
         */
        load(option?: Visio.Interfaces.ShapeCollectionLoadOptions & Visio.Interfaces.CollectionLoadOptions): Visio.ShapeCollection;
        load(option?: string | string[]): Visio.ShapeCollection;
        load(option?: OfficeExtension.LoadOption): Visio.ShapeCollection;
        toJSON(): Visio.Interfaces.ShapeCollectionData;
    }
    /**
     *
     * Represents the Shape class.
     *
     * [Api set:  1.1]
     */
    export class Shape extends OfficeExtension.ClientObject {
        /**
         *
         * Returns the Comments Collection
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
         * Gets SubShape Collection.
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
         * Shape's Identifier.
         *
         * [Api set:  1.1]
         */
        readonly id: number;
        /**
         *
         * Shape's name.
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
         * Shape's Text.
         *
         * [Api set:  1.1]
         */
        readonly text: string;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.ShapeUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Shape): void;
        /**
         *
         * Returns the BoundingBox object that specifies bounding box of the shape.
         *
         * [Api set:  1.1]
         */
        getBounds(): OfficeExtension.ClientResult<Visio.BoundingBox>;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: Visio.Interfaces.ShapeLoadOptions): Visio.Shape;
        load(option?: string | string[]): Visio.Shape;
        load(option?: {
            select?: string;
            expand?: string;
        }): Visio.Shape;
        toJSON(): Visio.Interfaces.ShapeData;
    }
    /**
     *
     * Represents the ShapeView class.
     *
     * [Api set:  1.1]
     */
    export class ShapeView extends OfficeExtension.ClientObject {
        /**
         *
         * Represents the highlight around the shape.
         *
         * [Api set:  1.1]
         */
        highlight: Visio.Highlight;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.ShapeViewUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: ShapeView): void;
        /**
         *
         * Adds an overlay on top of the shape.
         *
         * [Api set:  1.1]
         *
         * @param OverlayType - An Overlay Type -Text, Image.
         * @param Content - Content of Overlay.
         * @param OverlayHorizontalAlignment - Horizontal Alignment of Overlay - Left, Center, Right
         * @param OverlayVerticalAlignment - Vertical Alignment of Overlay - Top, Middle, Bottom
         * @param Width - Overlay Width.
         * @param Height - Overlay Height.
         */
        addOverlay(OverlayType: string, Content: string, OverlayHorizontalAlignment: string, OverlayVerticalAlignment: string, Width: number, Height: number): OfficeExtension.ClientResult<number>;
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
         */
        load(option?: Visio.Interfaces.ShapeViewLoadOptions): Visio.ShapeView;
        load(option?: string | string[]): Visio.ShapeView;
        load(option?: {
            select?: string;
            expand?: string;
        }): Visio.ShapeView;
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
        /** Gets the loaded child items in this collection. */
        readonly items: Array<Visio.ShapeDataItem>;
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
         */
        load(option?: Visio.Interfaces.ShapeDataItemCollectionLoadOptions & Visio.Interfaces.CollectionLoadOptions): Visio.ShapeDataItemCollection;
        load(option?: string | string[]): Visio.ShapeDataItemCollection;
        load(option?: OfficeExtension.LoadOption): Visio.ShapeDataItemCollection;
        toJSON(): Visio.Interfaces.ShapeDataItemCollectionData;
    }
    /**
     *
     * Represents the ShapeDataItem.
     *
     * [Api set:  1.1]
     */
    export class ShapeDataItem extends OfficeExtension.ClientObject {
        /**
         *
         * A string that specifies the format of the shape data item.
         *
         * [Api set:  1.1]
         */
        readonly format: string;
        /**
         *
         * A string that specifies the formatted value of the shape data item.
         *
         * [Api set:  1.1]
         */
        readonly formattedValue: string;
        /**
         *
         * A string that specifies the label of the shape data item.
         *
         * [Api set:  1.1]
         */
        readonly label: string;
        /**
         *
         * A string that specifies the value of the shape data item.
         *
         * [Api set:  1.1]
         */
        readonly value: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: Visio.Interfaces.ShapeDataItemLoadOptions): Visio.ShapeDataItem;
        load(option?: string | string[]): Visio.ShapeDataItem;
        load(option?: {
            select?: string;
            expand?: string;
        }): Visio.ShapeDataItem;
        toJSON(): Visio.Interfaces.ShapeDataItemData;
    }
    /**
     *
     * Represents the Hyperlink Collection.
     *
     * [Api set:  1.1]
     */
    export class HyperlinkCollection extends OfficeExtension.ClientObject {
        /** Gets the loaded child items in this collection. */
        readonly items: Array<Visio.Hyperlink>;
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
         */
        load(option?: Visio.Interfaces.HyperlinkCollectionLoadOptions & Visio.Interfaces.CollectionLoadOptions): Visio.HyperlinkCollection;
        load(option?: string | string[]): Visio.HyperlinkCollection;
        load(option?: OfficeExtension.LoadOption): Visio.HyperlinkCollection;
        toJSON(): Visio.Interfaces.HyperlinkCollectionData;
    }
    /**
     *
     * Represents the Hyperlink.
     *
     * [Api set:  1.1]
     */
    export class Hyperlink extends OfficeExtension.ClientObject {
        /**
         *
         * Gets the address of the Hyperlink object.
         *
         * [Api set:  1.1]
         */
        readonly address: string;
        /**
         *
         * Gets the description of a hyperlink.
         *
         * [Api set:  1.1]
         */
        readonly description: string;
        /**
         *
         * Gets the extra info of a hyperlink.
         *
         * [Api set:  1.1]
         */
        readonly extraInfo: string;
        /**
         *
         * Gets the sub-address of the Hyperlink object.
         *
         * [Api set:  1.1]
         */
        readonly subAddress: string;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: Visio.Interfaces.HyperlinkLoadOptions): Visio.Hyperlink;
        load(option?: string | string[]): Visio.Hyperlink;
        load(option?: {
            select?: string;
            expand?: string;
        }): Visio.Hyperlink;
        toJSON(): Visio.Interfaces.HyperlinkData;
    }
    /**
     *
     * Represents the CommentCollection for a given Shape.
     *
     * [Api set:  1.1]
     */
    export class CommentCollection extends OfficeExtension.ClientObject {
        /** Gets the loaded child items in this collection. */
        readonly items: Array<Visio.Comment>;
        /**
         *
         * Gets the number of Shape Data Items.
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
         */
        load(option?: Visio.Interfaces.CommentCollectionLoadOptions & Visio.Interfaces.CollectionLoadOptions): Visio.CommentCollection;
        load(option?: string | string[]): Visio.CommentCollection;
        load(option?: OfficeExtension.LoadOption): Visio.CommentCollection;
        toJSON(): Visio.Interfaces.CommentCollectionData;
    }
    /**
     *
     * Represents the Comment.
     *
     * [Api set:  1.1]
     */
    export class Comment extends OfficeExtension.ClientObject {
        /**
         *
         * A string that specifies the label of the shape data item.
         *
         * [Api set:  1.1]
         */
        author: string;
        /**
         *
         * A string that specifies the format of the shape data item.
         *
         * [Api set:  1.1]
         */
        date: string;
        /**
         *
         * A string that specifies the value of the shape data item.
         *
         * [Api set:  1.1]
         */
        text: string;
        /** Sets multiple properties on the object at the same time, based on JSON input. */
        set(properties: Interfaces.CommentUpdateData, options?: {
            /**
             * Throw an error if the passed-in property list includes read-only properties (default = true).
             */
            throwOnReadOnly?: boolean;
        }): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Comment): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: Visio.Interfaces.CommentLoadOptions): Visio.Comment;
        load(option?: string | string[]): Visio.Comment;
        load(option?: {
            select?: string;
            expand?: string;
        }): Visio.Comment;
        toJSON(): Visio.Interfaces.CommentData;
    }
    /**
     *
     * Represents the Selection in the page.
     *
     * [Api set:  1.1]
     */
    export class Selection extends OfficeExtension.ClientObject {
        /**
         *
         * Gets the Shapes of the Selection
         *
         * [Api set:  1.1]
         */
        readonly shapes: Visio.ShapeCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[]): Visio.Selection;
        load(option?: {
            select?: string;
            expand?: string;
        }): Visio.Selection;
        toJSON(): Visio.Interfaces.SelectionData;
    }
    /**
     *
     * Represents the Horizontal Alignment of the Overlay relative to the shape.
     *
     * [Api set:  1.1]
     */
    export namespace OverlayHorizontalAlignment {
        /**
         *
         * left
         *
         */
        var left: string;
        /**
         *
         * center
         *
         */
        var center: string;
        /**
         *
         * right
         *
         */
        var right: string;
    }
    /**
     *
     * Represents the Vertical Alignment of the Overlay relative to the shape.
     *
     * [Api set:  1.1]
     */
    export namespace OverlayVerticalAlignment {
        /**
         *
         * top
         *
         */
        var top: string;
        /**
         *
         * middle
         *
         */
        var middle: string;
        /**
         *
         * bottom
         *
         */
        var bottom: string;
    }
    /**
     *
     * Represents the type of the overlay.
     *
     * [Api set:  1.1]
     */
    export namespace OverlayType {
        /**
         *
         * text
         *
         */
        var text: string;
        /**
         *
         * image
         *
         */
        var image: string;
    }
    /**
     *
     * Toolbar IDs of the app
     *
     * [Api set:  1.1]
     */
    export namespace ToolBarType {
        /**
         *
         * CommandBar
         *
         */
        var commandBar: string;
        /**
         *
         * PageNavigationBar
         *
         */
        var pageNavigationBar: string;
        /**
         *
         * StatusBar
         *
         */
        var statusBar: string;
    }
    export namespace ErrorCodes {
        var accessDenied: string;
        var generalException: string;
        var invalidArgument: string;
        var itemNotFound: string;
        var notImplemented: string;
        var unsupportedOperation: string;
    }
    export module Interfaces {
        export interface CollectionLoadOptions {
            $top?: number;
            $skip?: number;
        }
        /** An interface for updating data on the Application object, for use in "application.set({ ... })". */
        export interface ApplicationUpdateData {
            /**
             *
             * Show/Hide the application borders.
             *
             * [Api set:  1.1]
             */
            showBorders?: boolean;
            /**
             *
             * Show or Hide the standard toolbars.
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
             * Disable Zoom.
             *
             * [Api set:  1.1]
             */
            disableZoom?: boolean;
            /**
             *
             * Disable Hyperlinks.
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
             * Get/Set Page's Zoom level. The value can be between 10 and 400 and denotes the percentage of zoom.
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
             * A string that specifies the label of the shape data item.
             *
             * [Api set:  1.1]
             */
            author?: string;
            /**
             *
             * A string that specifies the format of the shape data item.
             *
             * [Api set:  1.1]
             */
            date?: string;
            /**
             *
             * A string that specifies the value of the shape data item.
             *
             * [Api set:  1.1]
             */
            text?: string;
        }
        /** An interface describing the data returned by calling "application.toJSON()". */
        export interface ApplicationData {
            /**
             *
             * Show/Hide the application borders.
             *
             * [Api set:  1.1]
             */
            showBorders?: boolean;
            /**
             *
             * Show or Hide the standard toolbars.
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
            * Returns the DocumentView object.
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
             * Disable Zoom.
             *
             * [Api set:  1.1]
             */
            disableZoom?: boolean;
            /**
             *
             * Disable Hyperlinks.
             *
             * [Api set:  1.1]
             */
            hideDiagramBoundary?: boolean;
        }
        /** An interface describing the data returned by calling "page.toJSON()". */
        export interface PageData {
            /**
            *
            * All shapes in the page. Read-only.
            *
            * [Api set:  1.1]
            */
            allShapes?: Visio.Interfaces.ShapeData[];
            /**
            *
            * Returns the Comments Collection
            *
            * [Api set:  1.1]
            */
            comments?: Visio.Interfaces.CommentData[];
            /**
            *
            * Shapes at root level, in the page. Read-only.
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
             * Index of the Page.
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
             * Get/Set Page's Zoom level. The value can be between 10 and 400 and denotes the percentage of zoom.
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
            * Returns the Comments Collection
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
            * Gets SubShape Collection.
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
             * Shape's Identifier.
             *
             * [Api set:  1.1]
             */
            id?: number;
            /**
             *
             * Shape's name.
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
             * Shape's Text.
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
             * A string that specifies the format of the shape data item.
             *
             * [Api set:  1.1]
             */
            format?: string;
            /**
             *
             * A string that specifies the formatted value of the shape data item.
             *
             * [Api set:  1.1]
             */
            formattedValue?: string;
            /**
             *
             * A string that specifies the label of the shape data item.
             *
             * [Api set:  1.1]
             */
            label?: string;
            /**
             *
             * A string that specifies the value of the shape data item.
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
             * Gets the address of the Hyperlink object.
             *
             * [Api set:  1.1]
             */
            address?: string;
            /**
             *
             * Gets the description of a hyperlink.
             *
             * [Api set:  1.1]
             */
            description?: string;
            /**
             *
             * Gets the extra info of a hyperlink.
             *
             * [Api set:  1.1]
             */
            extraInfo?: string;
            /**
             *
             * Gets the sub-address of the Hyperlink object.
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
             * A string that specifies the label of the shape data item.
             *
             * [Api set:  1.1]
             */
            author?: string;
            /**
             *
             * A string that specifies the format of the shape data item.
             *
             * [Api set:  1.1]
             */
            date?: string;
            /**
             *
             * A string that specifies the value of the shape data item.
             *
             * [Api set:  1.1]
             */
            text?: string;
        }
        /** An interface describing the data returned by calling "selection.toJSON()". */
        export interface SelectionData {
            /**
            *
            * Gets the Shapes of the Selection
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
             * Show/Hide the application borders.
             *
             * [Api set:  1.1]
             */
            showBorders?: boolean;
            /**
             *
             * Show or Hide the standard toolbars.
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
             * Disable Zoom.
             *
             * [Api set:  1.1]
             */
            disableZoom?: boolean;
            /**
             *
             * Disable Hyperlinks.
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
             * Index of the Page.
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
             * Get/Set Page's Zoom level. The value can be between 10 and 400 and denotes the percentage of zoom.
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
             * For EACH ITEM in the collection: Index of the Page.
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
             * For EACH ITEM in the collection: Shape's Identifier.
             *
             * [Api set:  1.1]
             */
            id?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Shape's name.
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
             * For EACH ITEM in the collection: Shape's Text.
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
             * Shape's Identifier.
             *
             * [Api set:  1.1]
             */
            id?: boolean;
            /**
             *
             * Shape's name.
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
             * Shape's Text.
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
             * For EACH ITEM in the collection: A string that specifies the format of the shape data item.
             *
             * [Api set:  1.1]
             */
            format?: boolean;
            /**
             *
             * For EACH ITEM in the collection: A string that specifies the formatted value of the shape data item.
             *
             * [Api set:  1.1]
             */
            formattedValue?: boolean;
            /**
             *
             * For EACH ITEM in the collection: A string that specifies the label of the shape data item.
             *
             * [Api set:  1.1]
             */
            label?: boolean;
            /**
             *
             * For EACH ITEM in the collection: A string that specifies the value of the shape data item.
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
             * A string that specifies the format of the shape data item.
             *
             * [Api set:  1.1]
             */
            format?: boolean;
            /**
             *
             * A string that specifies the formatted value of the shape data item.
             *
             * [Api set:  1.1]
             */
            formattedValue?: boolean;
            /**
             *
             * A string that specifies the label of the shape data item.
             *
             * [Api set:  1.1]
             */
            label?: boolean;
            /**
             *
             * A string that specifies the value of the shape data item.
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
             * For EACH ITEM in the collection: Gets the address of the Hyperlink object.
             *
             * [Api set:  1.1]
             */
            address?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the description of a hyperlink.
             *
             * [Api set:  1.1]
             */
            description?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the extra info of a hyperlink.
             *
             * [Api set:  1.1]
             */
            extraInfo?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the sub-address of the Hyperlink object.
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
             * Gets the address of the Hyperlink object.
             *
             * [Api set:  1.1]
             */
            address?: boolean;
            /**
             *
             * Gets the description of a hyperlink.
             *
             * [Api set:  1.1]
             */
            description?: boolean;
            /**
             *
             * Gets the extra info of a hyperlink.
             *
             * [Api set:  1.1]
             */
            extraInfo?: boolean;
            /**
             *
             * Gets the sub-address of the Hyperlink object.
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
             * For EACH ITEM in the collection: A string that specifies the label of the shape data item.
             *
             * [Api set:  1.1]
             */
            author?: boolean;
            /**
             *
             * For EACH ITEM in the collection: A string that specifies the format of the shape data item.
             *
             * [Api set:  1.1]
             */
            date?: boolean;
            /**
             *
             * For EACH ITEM in the collection: A string that specifies the value of the shape data item.
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
             * A string that specifies the label of the shape data item.
             *
             * [Api set:  1.1]
             */
            author?: boolean;
            /**
             *
             * A string that specifies the format of the shape data item.
             *
             * [Api set:  1.1]
             */
            date?: boolean;
            /**
             *
             * A string that specifies the value of the shape data item.
             *
             * [Api set:  1.1]
             */
            text?: boolean;
        }
    }
}
declare module Visio {
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
    export function run<T>(batch: (context: Visio.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
    /**
     * Executes a batch script that performs actions on the Visio object model, using the request context of a previously-created API object.
     * @param object - A previously-created API object. The batch will use the same request context as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
     * @param batch - A function that takes in an Visio.RequestContext and returns a promise (typically, just the result of "context.sync()"). When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     */
    export function run<T>(object: OfficeExtension.ClientObject | OfficeExtension.EmbeddedSession, batch: (context: Visio.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
    /**
     * Executes a batch script that performs actions on the Visio object model, using the RequestContext of a previously-created object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param contextObject - A previously-created Visio.RequestContext. This context will get re-used by the batch function (instead of having a new context created). This means that the batch will be able to pick up changes made to existing API objects, if those objects were derived from this same context.
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Visio application. Since the Office add-in and the Visio application run in two different processes, the RequestContext is required to get access to the Visio object model from the add-in.
     */
    export function run<T>(contextObject: OfficeExtension.ClientRequestContext, batch: (context: Visio.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
    /**
     * Executes a batch script that performs actions on the Visio object model, using the request context of previously-created API objects.
     * @param objects - An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared request context, which means that any changes applied to these objects will be picked up by "context.sync()".
     * @param batch - A function that takes in a Visio.RequestContext and returns a promise (typically, just the result of "context.sync()"). When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     */
    export function run<T>(objects: OfficeExtension.ClientObject[], batch: (context: Visio.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
}



////////////////////////////////////////////////////////////////
//////////////////////// End Visio APIs ////////////////////////
////////////////////////////////////////////////////////////////