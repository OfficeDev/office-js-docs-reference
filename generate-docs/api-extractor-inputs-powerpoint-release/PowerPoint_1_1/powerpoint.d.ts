import { OfficeExtension } from "../../api-extractor-inputs-office/office"
import { Office as Outlook} from "../../api-extractor-inputs-outlook/outlook"
////////////////////////////////////////////////////////////////
//////////////////// Begin PowerPoint APIs /////////////////////
////////////////////////////////////////////////////////////////

export declare namespace PowerPoint {
    /**
     * @remarks
     * [Api set: PowerPointApi 1.0]
     */
    export class Application extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Create a new instance of the `PowerPoint.Application` object.
         */
        static newObject(context: OfficeExtension.ClientRequestContext): PowerPoint.Application;
        /**
         * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
         * Whereas the original `PowerPoint.Application` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.ApplicationData`) that contains shallow copies of any loaded child properties from the original object.
         */
        toJSON(): {
            [key: string]: string;
        };
    }
    /**
     * @remarks
     * [Api set: PowerPointApi 1.0]
     */
    export class Presentation extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        
        
        
        
        
        readonly title: string;
        
        
        
        
        
        
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
         * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
         * Whereas the original `PowerPoint.Presentation` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.PresentationData`) that contains shallow copies of any loaded child properties from the original object.
         */
        toJSON(): PowerPoint.Interfaces.PresentationData;
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    enum ErrorCodes {
        generalException = "GeneralException",
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
            * Specify the number of items in the collection that are to be skipped and not included in the result. If top is specified, the selection of result will start after skipping the specified number of items.
            */
            $skip?: number;
        }
        /** An interface for updating data on the `CustomXmlPartScopedCollection` object, for use in `customXmlPartScopedCollection.set({ ... })`. */
        export interface CustomXmlPartScopedCollectionUpdateData {
            items?: PowerPoint.Interfaces.CustomXmlPartData[];
        }
        /** An interface for updating data on the `CustomXmlPartCollection` object, for use in `customXmlPartCollection.set({ ... })`. */
        export interface CustomXmlPartCollectionUpdateData {
            items?: PowerPoint.Interfaces.CustomXmlPartData[];
        }
        /** An interface for updating data on the `Hyperlink` object, for use in `hyperlink.set({ ... })`. */
        export interface HyperlinkUpdateData {
            
            
        }
        /** An interface for updating data on the `HyperlinkCollection` object, for use in `hyperlinkCollection.set({ ... })`. */
        export interface HyperlinkCollectionUpdateData {
            items?: PowerPoint.Interfaces.HyperlinkData[];
        }
        /** An interface for updating data on the `ShapeCollection` object, for use in `shapeCollection.set({ ... })`. */
        export interface ShapeCollectionUpdateData {
            items?: PowerPoint.Interfaces.ShapeData[];
        }
        /** An interface for updating data on the `SlideLayoutCollection` object, for use in `slideLayoutCollection.set({ ... })`. */
        export interface SlideLayoutCollectionUpdateData {
            items?: PowerPoint.Interfaces.SlideLayoutData[];
        }
        /** An interface for updating data on the `Tag` object, for use in `tag.set({ ... })`. */
        export interface TagUpdateData {
            
        }
        /** An interface for updating data on the `TagCollection` object, for use in `tagCollection.set({ ... })`. */
        export interface TagCollectionUpdateData {
            items?: PowerPoint.Interfaces.TagData[];
        }
        /** An interface for updating data on the `ShapeFill` object, for use in `shapeFill.set({ ... })`. */
        export interface ShapeFillUpdateData {
            
            
        }
        /** An interface for updating data on the `ShapeScopedCollection` object, for use in `shapeScopedCollection.set({ ... })`. */
        export interface ShapeScopedCollectionUpdateData {
            items?: PowerPoint.Interfaces.ShapeData[];
        }
        /** An interface for updating data on the `ShapeLineFormat` object, for use in `shapeLineFormat.set({ ... })`. */
        export interface ShapeLineFormatUpdateData {
            
            
            
            
            
            
        }
        /** An interface for updating data on the `BulletFormat` object, for use in `bulletFormat.set({ ... })`. */
        export interface BulletFormatUpdateData {
            
        }
        /** An interface for updating data on the `ParagraphFormat` object, for use in `paragraphFormat.set({ ... })`. */
        export interface ParagraphFormatUpdateData {
            
        }
        /** An interface for updating data on the `ShapeFont` object, for use in `shapeFont.set({ ... })`. */
        export interface ShapeFontUpdateData {
            
            
            
            
            
            
        }
        /** An interface for updating data on the `TextRange` object, for use in `textRange.set({ ... })`. */
        export interface TextRangeUpdateData {
            
            
            
        }
        /** An interface for updating data on the `TextFrame` object, for use in `textFrame.set({ ... })`. */
        export interface TextFrameUpdateData {
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `Shape` object, for use in `shape.set({ ... })`. */
        export interface ShapeUpdateData {
            
            
            
            
            
        }
        /** An interface for updating data on the `CustomProperty` object, for use in `customProperty.set({ ... })`. */
        export interface CustomPropertyUpdateData {
            
        }
        /** An interface for updating data on the `CustomPropertyCollection` object, for use in `customPropertyCollection.set({ ... })`. */
        export interface CustomPropertyCollectionUpdateData {
            items?: PowerPoint.Interfaces.CustomPropertyData[];
        }
        /** An interface for updating data on the `DocumentProperties` object, for use in `documentProperties.set({ ... })`. */
        export interface DocumentPropertiesUpdateData {
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `SlideCollection` object, for use in `slideCollection.set({ ... })`. */
        export interface SlideCollectionUpdateData {
            items?: PowerPoint.Interfaces.SlideData[];
        }
        /** An interface for updating data on the `SlideScopedCollection` object, for use in `slideScopedCollection.set({ ... })`. */
        export interface SlideScopedCollectionUpdateData {
            items?: PowerPoint.Interfaces.SlideData[];
        }
        /** An interface for updating data on the `SlideMasterCollection` object, for use in `slideMasterCollection.set({ ... })`. */
        export interface SlideMasterCollectionUpdateData {
            items?: PowerPoint.Interfaces.SlideMasterData[];
        }
        /** An interface describing the data returned by calling `presentation.toJSON()`. */
        export interface PresentationData {
            
            title?: string;
        }
        /** An interface describing the data returned by calling `customXmlPart.toJSON()`. */
        export interface CustomXmlPartData {
            
            
        }
        /** An interface describing the data returned by calling `customXmlPartScopedCollection.toJSON()`. */
        export interface CustomXmlPartScopedCollectionData {
            items?: PowerPoint.Interfaces.CustomXmlPartData[];
        }
        /** An interface describing the data returned by calling `customXmlPartCollection.toJSON()`. */
        export interface CustomXmlPartCollectionData {
            items?: PowerPoint.Interfaces.CustomXmlPartData[];
        }
        /** An interface describing the data returned by calling `hyperlink.toJSON()`. */
        export interface HyperlinkData {
            
            
        }
        /** An interface describing the data returned by calling `hyperlinkCollection.toJSON()`. */
        export interface HyperlinkCollectionData {
            items?: PowerPoint.Interfaces.HyperlinkData[];
        }
        /** An interface describing the data returned by calling `shapeCollection.toJSON()`. */
        export interface ShapeCollectionData {
            items?: PowerPoint.Interfaces.ShapeData[];
        }
        /** An interface describing the data returned by calling `slideLayout.toJSON()`. */
        export interface SlideLayoutData {
            
            
        }
        /** An interface describing the data returned by calling `slideLayoutCollection.toJSON()`. */
        export interface SlideLayoutCollectionData {
            items?: PowerPoint.Interfaces.SlideLayoutData[];
        }
        /** An interface describing the data returned by calling `slideMaster.toJSON()`. */
        export interface SlideMasterData {
            
            
        }
        /** An interface describing the data returned by calling `tag.toJSON()`. */
        export interface TagData {
            
            
        }
        /** An interface describing the data returned by calling `tagCollection.toJSON()`. */
        export interface TagCollectionData {
            items?: PowerPoint.Interfaces.TagData[];
        }
        /** An interface describing the data returned by calling `slide.toJSON()`. */
        export interface SlideData {
            
        }
        /** An interface describing the data returned by calling `shapeFill.toJSON()`. */
        export interface ShapeFillData {
            
            
            
        }
        /** An interface describing the data returned by calling `shapeScopedCollection.toJSON()`. */
        export interface ShapeScopedCollectionData {
            items?: PowerPoint.Interfaces.ShapeData[];
        }
        /** An interface describing the data returned by calling `shapeLineFormat.toJSON()`. */
        export interface ShapeLineFormatData {
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `bulletFormat.toJSON()`. */
        export interface BulletFormatData {
            
        }
        /** An interface describing the data returned by calling `paragraphFormat.toJSON()`. */
        export interface ParagraphFormatData {
            
        }
        /** An interface describing the data returned by calling `shapeFont.toJSON()`. */
        export interface ShapeFontData {
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `textRange.toJSON()`. */
        export interface TextRangeData {
            
            
            
        }
        /** An interface describing the data returned by calling `textFrame.toJSON()`. */
        export interface TextFrameData {
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `shape.toJSON()`. */
        export interface ShapeData {
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `customProperty.toJSON()`. */
        export interface CustomPropertyData {
            
            
            
        }
        /** An interface describing the data returned by calling `customPropertyCollection.toJSON()`. */
        export interface CustomPropertyCollectionData {
            items?: PowerPoint.Interfaces.CustomPropertyData[];
        }
        /** An interface describing the data returned by calling `documentProperties.toJSON()`. */
        export interface DocumentPropertiesData {
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `slideCollection.toJSON()`. */
        export interface SlideCollectionData {
            items?: PowerPoint.Interfaces.SlideData[];
        }
        /** An interface describing the data returned by calling `slideScopedCollection.toJSON()`. */
        export interface SlideScopedCollectionData {
            items?: PowerPoint.Interfaces.SlideData[];
        }
        /** An interface describing the data returned by calling `slideMasterCollection.toJSON()`. */
        export interface SlideMasterCollectionData {
            items?: PowerPoint.Interfaces.SlideMasterData[];
        }
        /**
         * @remarks
         * [Api set: PowerPointApi 1.0]
         */
        export interface PresentationLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            
            title?: boolean;
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
    export function run<T>(batch: (context: PowerPoint.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
    /**
     * Executes a batch script that performs actions on the PowerPoint object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param object - A previously-created API object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the PowerPoint application. Since the Office add-in and the PowerPoint application run in two different processes, the RequestContext is required to get access to the PowerPoint object model from the add-in.
     */
    export function run<T>(object: OfficeExtension.ClientObject, batch: (context: PowerPoint.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
    /**
     * Executes a batch script that performs actions on the PowerPoint object model, using the RequestContext of previously-created API objects.
     * @param objects - An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the PowerPoint application. Since the Office add-in and the PowerPoint application run in two different processes, the RequestContext is required to get access to the PowerPoint object model from the add-in.
     */
    export function run<T>(objects: OfficeExtension.ClientObject[], batch: (context: PowerPoint.RequestContext) => OfficeExtension.IPromise<T>): OfficeExtension.IPromise<T>;
}
export declare namespace PowerPoint {
    /**
     * Creates and opens a new presentation. Optionally, the presentation can be pre-populated with a Base64-encoded .pptx file.
     *
     * [Api set: PowerPointApi 1.1]
     *
     * @param base64File - Optional. The Base64-encoded .pptx file. The default value is null.
     */
    export function createPresentation(base64File?: string): Promise<void>;
}


////////////////////////////////////////////////////////////////
///////////////////// End PowerPoint APIs //////////////////////
////////////////////////////////////////////////////////////////