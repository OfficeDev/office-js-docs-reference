import { OfficeExtension } from "../../api-extractor-inputs-office/office"
import { Office as Outlook} from "../../api-extractor-inputs-outlook/outlook"
////////////////////////////////////////////////////////////////
//////////////////// Begin PowerPoint APIs /////////////////////
////////////////////////////////////////////////////////////////

export declare namespace PowerPoint {
    /**
     * [Api set: PowerPointApi 1.0]
     */
    export class Application extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Create a new instance of PowerPoint.Application object
         */
        static newObject(context: OfficeExtension.ClientRequestContext): PowerPoint.Application;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original PowerPoint.Application object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.ApplicationData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): {
            [key: string]: string;
        };
    }
    /**
     * [Api set: PowerPointApi 1.0]
     */
    export class Presentation extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         *
         * Returns an ordered collection of slides in the presentation.
         *
         * [Api set: PowerPointApi 1.2]
         */
        readonly slides: PowerPoint.SlideCollection;
        readonly title: string;
        /**
         * Inserts the specified slides from a presentation into the current presentation.
         *
         * [Api set: PowerPointApi 1.2]
         *
         * @param base64File - The base64-encoded string representing the source presentation file.
         * @param options - The options that define which slides will be inserted, where the new slides will go, and which presentation's formatting will be used.
         */
        insertSlidesFromBase64(base64File: string, options?: PowerPoint.InsertSlideOptions): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.PresentationLoadOptions): PowerPoint.Presentation;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.Presentation;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): PowerPoint.Presentation;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original PowerPoint.Presentation object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.PresentationData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): PowerPoint.Interfaces.PresentationData;
    }
    /**
     *
     * Specifies the formatting options for when slides are inserted.
     *
     * [Api set: PowerPointApi 1.2]
     */
    enum InsertSlideFormatting {
        /**
         * Copy the source theme into the target presentation and use that theme.
         *
         */
        keepSourceFormatting = "KeepSourceFormatting",
        /**
         * Use the existing theme in the target presentation.
         *
         */
        useDestinationTheme = "UseDestinationTheme",
    }
    /**
     *
     * Represents the available options when inserting slides.
     *
     * [Api set: PowerPointApi 1.2]
     */
    export interface InsertSlideOptions {
        /**
         *
         * Specifies which formatting to use during slide insertion.
                    The default option is to use "KeepSourceFormatting".
         *
         * [Api set: PowerPointApi 1.2]
         */
        formatting?: PowerPoint.InsertSlideFormatting | "KeepSourceFormatting" | "UseDestinationTheme";
        /**
         *
         * Specifies the slides from the source presentation that will be inserted into the current presentation. These slides are represented by their IDs which can be retrieved from a `Slide` object.
                    The order of these slides is preserved during the insertion.
                    If any of the source slides are not found, or if the IDs are invalid, the operation throws a `SlideNotFound` exception and no slides will be inserted.
                    All of the source slides will be inserted when `sourceSlideIds` is not provided (this is the default behavior).
         *
         * [Api set: PowerPointApi 1.2]
         */
        sourceSlideIds?: string[];
        /**
         *
         * Specifies where in the presentation the new slides will be inserted. The new slides will be inserted after the slide with the given slide ID.
                    If `targetSlideId` is not provided, the slides will be inserted at the beginning of the presentation.
                    If `targetSlideId` is invalid or if it is pointing to a non-existing slide, the operation throws a `SlideNotFound` exception and no slides will be inserted.
         *
         * [Api set: PowerPointApi 1.2]
         */
        targetSlideId?: string;
    }
    /**
     *
     * Represents a single slide of a presentation.
     *
     * [Api set: PowerPointApi 1.2]
     */
    export class Slide extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         *
         * Gets the unique ID of the slide.
         *
         * [Api set: PowerPointApi 1.2]
         */
        readonly id: string;
        /**
         * Deletes the slide from the presentation. Does nothing if the slide does not exist.
         *
         * [Api set: PowerPointApi 1.2]
         */
        delete(): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.SlideLoadOptions): PowerPoint.Slide;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.Slide;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): PowerPoint.Slide;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original PowerPoint.Slide object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.SlideData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): PowerPoint.Interfaces.SlideData;
    }
    /**
     *
     * Represents the collection of slides in the presentation.
     *
     * [Api set: PowerPointApi 1.2]
     */
    export class SlideCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: PowerPoint.Slide[];
        /**
         * Gets the number of slides in the collection.
         *
         * [Api set: PowerPointApi 1.2]
         * @returns The number of slides in the collection.
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         * Gets a slide using its unique ID.
         *
         * [Api set: PowerPointApi 1.2]
         *
         * @param key - The ID of the slide.
         * @returns The slide with the unique ID. If such a slide does not exist, an error is thrown.
         */
        getItem(key: string): PowerPoint.Slide;
        /**
         * Gets a slide using its zero-based index in the collection. Slides are stored in the same order as they
                    are shown in the presentation.
         *
         * [Api set: PowerPointApi 1.2]
         *
         * @param index - The index of the slide in the collection.
         * @returns The slide at the given index. An error is thrown if index is out of range.
         */
        getItemAt(index: number): PowerPoint.Slide;
        /**
         * Gets a slide using its unique ID. If such a slide does not exist, an object with an `isNullObject` property set to true is returned. For further information,
                    see {@link https://docs.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties | *OrNullObject methods
                    and properties}.
         *
         * [Api set: PowerPointApi 1.2]
         *
         * @param id - The ID of the slide.
         * @returns The slide with the unique ID.
         */
        getItemOrNullObject(id: string): PowerPoint.Slide;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: PowerPoint.Interfaces.SlideCollectionLoadOptions & PowerPoint.Interfaces.CollectionLoadOptions): PowerPoint.SlideCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): PowerPoint.SlideCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): PowerPoint.SlideCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `PowerPoint.SlideCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.SlideCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): PowerPoint.Interfaces.SlideCollectionData;
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
        /** An interface for updating data on the SlideCollection object, for use in `slideCollection.set({ ... })`. */
        export interface SlideCollectionUpdateData {
            items?: PowerPoint.Interfaces.SlideData[];
        }
        /** An interface describing the data returned by calling `presentation.toJSON()`. */
        export interface PresentationData {
            title?: string;
        }
        /** An interface describing the data returned by calling `slide.toJSON()`. */
        export interface SlideData {
            /**
             *
             * Gets the unique ID of the slide.
             *
             * [Api set: PowerPointApi 1.2]
             */
            id?: string;
        }
        /** An interface describing the data returned by calling `slideCollection.toJSON()`. */
        export interface SlideCollectionData {
            items?: PowerPoint.Interfaces.SlideData[];
        }
        /**
         * [Api set: PowerPointApi 1.0]
         */
        export interface PresentationLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            title?: boolean;
        }
        /**
         *
         * Represents a single slide of a presentation.
         *
         * [Api set: PowerPointApi 1.2]
         */
        export interface SlideLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             *
             * Gets the unique ID of the slide.
             *
             * [Api set: PowerPointApi 1.2]
             */
            id?: boolean;
        }
        /**
         *
         * Represents the collection of slides in the presentation.
         *
         * [Api set: PowerPointApi 1.2]
         */
        export interface SlideCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             *
             * For EACH ITEM in the collection: Gets the unique ID of the slide.
             *
             * [Api set: PowerPointApi 1.2]
             */
            id?: boolean;
        }
    }
}
export declare namespace PowerPoint {
    /**
     * The RequestContext object facilitates requests to the PowerPoint application. Since the Office add-in and the PowerPoint application run in two different processes, the request context is required to get access to the PowerPoint object model from the add-in.
     */
    export class RequestContext extends OfficeExtension.ClientRequestContext {
        constructor(url?: string);
        readonly presentation: Presentation;
        readonly application: Application;
    }
    /**
     * Executes a batch script that performs actions on the PowerPoint object model, using a new RequestContext. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the PowerPoint application. Since the Office add-in and the PowerPoint application run in two different processes, the RequestContext is required to get access to the PowerPoint object model from the add-in.
     */
    export function run<T>(batch: (context: PowerPoint.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the PowerPoint object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param object - A previously-created API object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the PowerPoint application. Since the Office add-in and the PowerPoint application run in two different processes, the RequestContext is required to get access to the PowerPoint object model from the add-in.
     */
    export function run<T>(object: OfficeExtension.ClientObject, batch: (context: PowerPoint.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the PowerPoint object model, using the RequestContext of previously-created API objects.
     * @param objects - An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the PowerPoint application. Since the Office add-in and the PowerPoint application run in two different processes, the RequestContext is required to get access to the PowerPoint object model from the add-in.
     */
    export function run<T>(objects: OfficeExtension.ClientObject[], batch: (context: PowerPoint.RequestContext) => Promise<T>): Promise<T>;
}
export declare namespace PowerPoint {
    /**
     * Creates and opens a new presentation. Optionally, the presentation can be pre-populated with a base64-encoded .pptx file.
     *
     * [Api set: PowerPointApi 1.1]
     *
     * @param base64File - Optional. The base64-encoded .pptx file. The default value is null.
     */
    export function createPresentation(base64File?: string): Promise<void>;
}


////////////////////////////////////////////////////////////////
///////////////////// End PowerPoint APIs //////////////////////
////////////////////////////////////////////////////////////////