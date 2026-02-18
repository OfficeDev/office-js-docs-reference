import { OfficeExtension } from "../../api-extractor-inputs-office/office"
import { Office as Outlook} from "../../api-extractor-inputs-outlook/outlook"
////////////////////////////////////////////////////////////////
//////////////////// Begin PowerPoint APIs /////////////////////
////////////////////////////////////////////////////////////////

export declare namespace PowerPoint {
    /**
     * Represents the PowerPoint application that manages the presentation.
     *
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
     * The `Presentation` object is the top-level object with one or more slides that contain the contents of the presentation.
                To learn more about the PowerPoint object model,
                see {@link https://learn.microsoft.com/office/dev/add-ins/powerpoint/core-concepts | PowerPoint JavaScript object model}.
     *
     * @remarks
     * [Api set: PowerPointApi 1.0]
     */
    export class Presentation extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        
        
        
        
        /**
         * Returns an ordered collection of slides in the presentation.
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         */
        readonly slides: PowerPoint.SlideCollection;
        
        
        /**
         * Returns the title of the presentation.
         *
         * @remarks
         * [Api set: PowerPointApi 1.0]
         */
        readonly title: string;
        
        
        
        
        /**
         * Inserts the specified slides from a presentation into the current presentation.
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         *
         * @param base64File - The Base64-encoded string representing the source presentation file.
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
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `PowerPoint.Presentation` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.PresentationData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): PowerPoint.Interfaces.PresentationData;
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    /**
     * Represents a single slide of a presentation.
     *
     * @remarks
     * [Api set: PowerPointApi 1.2]
     */
    export class Slide extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        
        
        
        
        
        
        
        /**
         * Gets the unique ID of the slide.
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         */
        readonly id: string;
        
        
        /**
         * Deletes the slide from the presentation. Does nothing if the slide doesn't exist.
         *
         * @remarks
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
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `PowerPoint.Slide` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.SlideData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): PowerPoint.Interfaces.SlideData;
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    /**
     * Specifies the formatting options for when slides are inserted.
     *
     * @remarks
     * [Api set: PowerPointApi 1.2]
     */
    enum InsertSlideFormatting {
        /**
         * Copy the source theme into the target presentation and use that theme.
         * @remarks
         * [Api set: PowerPointApi 1.2]
         */
        keepSourceFormatting = "KeepSourceFormatting",
        /**
         * Use the existing theme in the target presentation.
         * @remarks
         * [Api set: PowerPointApi 1.2]
         */
        useDestinationTheme = "UseDestinationTheme",
    }
    /**
     * Represents the available options when inserting slides.
     *
     * @remarks
     * [Api set: PowerPointApi 1.2]
     */
    export interface InsertSlideOptions {
        /**
         * Specifies which formatting to use during slide insertion.
                    The default option is to use "KeepSourceFormatting".
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         */
        formatting?: PowerPoint.InsertSlideFormatting | "KeepSourceFormatting" | "UseDestinationTheme";
        /**
         * Specifies the slides from the source presentation that will be inserted into the current presentation. These slides are represented by their IDs which can be retrieved from a `Slide` object.
                    The order of these slides is preserved during the insertion.
                    If any of the source slides aren't found, or if the IDs are invalid, the operation throws a `SlideNotFound` exception and no slides will be inserted.
                    All of the source slides will be inserted when `sourceSlideIds` isn't provided (this is the default behavior).
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         */
        sourceSlideIds?: string[];
        /**
         * Specifies where in the presentation the new slides will be inserted. The new slides will be inserted after the slide with the given slide ID.
                    If `targetSlideId` isn't provided, the slides will be inserted at the beginning of the presentation.
                    If `targetSlideId` is invalid or if it's pointing to a non-existing slide, the operation throws a `SlideNotFound` exception and no slides will be inserted.
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         */
        targetSlideId?: string;
    }
    
    /**
     * Represents the collection of slides in the presentation.
     *
     * @remarks
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
         * @remarks
         * [Api set: PowerPointApi 1.2]
         * @returns The number of slides in the collection.
         */
        getCount(): OfficeExtension.ClientResult<number>;
        /**
         * Gets a slide using its unique ID.
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         *
         * @param key - The ID of the slide.
         * @returns The slide with the unique ID. If such a slide doesn't exist, an error is thrown.
         */
        getItem(key: string): PowerPoint.Slide;
        /**
         * Gets a slide using its zero-based index in the collection. Slides are stored in the same order as they
                    are shown in the presentation.
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         *
         * @param index - The index of the slide in the collection.
         * @returns The slide at the given index. An error is thrown if index is out of range.
         */
        getItemAt(index: number): PowerPoint.Slide;
        /**
         * Gets a slide using its unique ID. If such a slide doesn't exist, an object with an `isNullObject` property set to true is returned. For further information, see
                    {@link https://learn.microsoft.com/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties | *OrNullObject methods and properties}.
         *
         * @remarks
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
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.)
        * Whereas the original `PowerPoint.SlideCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `PowerPoint.Interfaces.SlideCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): PowerPoint.Interfaces.SlideCollectionData;
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
            * Specify the number of items in the collection that are to be skipped and not included in the result. If `top` is specified, the selection of result will start after skipping the specified number of items.
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
        /** An interface for updating data on the `HyperlinkScopedCollection` object, for use in `hyperlinkScopedCollection.set({ ... })`. */
        export interface HyperlinkScopedCollectionUpdateData {
            items?: PowerPoint.Interfaces.HyperlinkData[];
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
        /** An interface for updating data on the `TextFrame` object, for use in `textFrame.set({ ... })`. */
        export interface TextFrameUpdateData {
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `TextRange` object, for use in `textRange.set({ ... })`. */
        export interface TextRangeUpdateData {
            
            
            
        }
        /** An interface for updating data on the `Hyperlink` object, for use in `hyperlink.set({ ... })`. */
        export interface HyperlinkUpdateData {
            
            
        }
        /** An interface for updating data on the `HyperlinkCollection` object, for use in `hyperlinkCollection.set({ ... })`. */
        export interface HyperlinkCollectionUpdateData {
            items?: PowerPoint.Interfaces.HyperlinkData[];
        }
        /** An interface for updating data on the `ShapeScopedCollection` object, for use in `shapeScopedCollection.set({ ... })`. */
        export interface ShapeScopedCollectionUpdateData {
            items?: PowerPoint.Interfaces.ShapeData[];
        }
        /** An interface for updating data on the `Border` object, for use in `border.set({ ... })`. */
        export interface BorderUpdateData {
            
            
            
            
        }
        /** An interface for updating data on the `Margins` object, for use in `margins.set({ ... })`. */
        export interface MarginsUpdateData {
            
            
            
            
        }
        /** An interface for updating data on the `ShapeFill` object, for use in `shapeFill.set({ ... })`. */
        export interface ShapeFillUpdateData {
            
            
        }
        /** An interface for updating data on the `TableCell` object, for use in `tableCell.set({ ... })`. */
        export interface TableCellUpdateData {
            
            
            
            
            
        }
        /** An interface for updating data on the `TableCellCollection` object, for use in `tableCellCollection.set({ ... })`. */
        export interface TableCellCollectionUpdateData {
            items?: PowerPoint.Interfaces.TableCellData[];
        }
        /** An interface for updating data on the `TableColumn` object, for use in `tableColumn.set({ ... })`. */
        export interface TableColumnUpdateData {
            
        }
        /** An interface for updating data on the `TableColumnCollection` object, for use in `tableColumnCollection.set({ ... })`. */
        export interface TableColumnCollectionUpdateData {
            items?: PowerPoint.Interfaces.TableColumnData[];
        }
        /** An interface for updating data on the `TableRow` object, for use in `tableRow.set({ ... })`. */
        export interface TableRowUpdateData {
            
        }
        /** An interface for updating data on the `TableRowCollection` object, for use in `tableRowCollection.set({ ... })`. */
        export interface TableRowCollectionUpdateData {
            items?: PowerPoint.Interfaces.TableRowData[];
        }
        /** An interface for updating data on the `TableStyleSettings` object, for use in `tableStyleSettings.set({ ... })`. */
        export interface TableStyleSettingsUpdateData {
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `ShapeCollection` object, for use in `shapeCollection.set({ ... })`. */
        export interface ShapeCollectionUpdateData {
            items?: PowerPoint.Interfaces.ShapeData[];
        }
        /** An interface for updating data on the `SlideBackgroundGradientFill` object, for use in `slideBackgroundGradientFill.set({ ... })`. */
        export interface SlideBackgroundGradientFillUpdateData {
            
        }
        /** An interface for updating data on the `SlideBackgroundPatternFill` object, for use in `slideBackgroundPatternFill.set({ ... })`. */
        export interface SlideBackgroundPatternFillUpdateData {
            
            
            
        }
        /** An interface for updating data on the `SlideBackgroundPictureOrTextureFill` object, for use in `slideBackgroundPictureOrTextureFill.set({ ... })`. */
        export interface SlideBackgroundPictureOrTextureFillUpdateData {
            
        }
        /** An interface for updating data on the `SlideBackgroundSolidFill` object, for use in `slideBackgroundSolidFill.set({ ... })`. */
        export interface SlideBackgroundSolidFillUpdateData {
            
            
        }
        /** An interface for updating data on the `SlideBackground` object, for use in `slideBackground.set({ ... })`. */
        export interface SlideBackgroundUpdateData {
            
            
        }
        /** An interface for updating data on the `SlideLayoutBackground` object, for use in `slideLayoutBackground.set({ ... })`. */
        export interface SlideLayoutBackgroundUpdateData {
            
            
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
        /** An interface for updating data on the `ShapeLineFormat` object, for use in `shapeLineFormat.set({ ... })`. */
        export interface ShapeLineFormatUpdateData {
            
            
            
            
            
            
        }
        /** An interface for updating data on the `Shape` object, for use in `shape.set({ ... })`. */
        export interface ShapeUpdateData {
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface for updating data on the `BindingCollection` object, for use in `bindingCollection.set({ ... })`. */
        export interface BindingCollectionUpdateData {
            items?: PowerPoint.Interfaces.BindingData[];
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
        /** An interface for updating data on the `PageSetup` object, for use in `pageSetup.set({ ... })`. */
        export interface PageSetupUpdateData {
            
            
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
            
            /**
             * Returns the title of the presentation.
             *
             * @remarks
             * [Api set: PowerPointApi 1.0]
             */
            title?: string;
        }
        /** An interface describing the data returned by calling `adjustments.toJSON()`. */
        export interface AdjustmentsData {
            
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
        /** An interface describing the data returned by calling `hyperlinkScopedCollection.toJSON()`. */
        export interface HyperlinkScopedCollectionData {
            items?: PowerPoint.Interfaces.HyperlinkData[];
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
        /** An interface describing the data returned by calling `textFrame.toJSON()`. */
        export interface TextFrameData {
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `textRange.toJSON()`. */
        export interface TextRangeData {
            
            
            
        }
        /** An interface describing the data returned by calling `hyperlink.toJSON()`. */
        export interface HyperlinkData {
            
            
            
        }
        /** An interface describing the data returned by calling `placeholderFormat.toJSON()`. */
        export interface PlaceholderFormatData {
            
            
        }
        /** An interface describing the data returned by calling `hyperlinkCollection.toJSON()`. */
        export interface HyperlinkCollectionData {
            items?: PowerPoint.Interfaces.HyperlinkData[];
        }
        /** An interface describing the data returned by calling `shapeScopedCollection.toJSON()`. */
        export interface ShapeScopedCollectionData {
            items?: PowerPoint.Interfaces.ShapeData[];
        }
        /** An interface describing the data returned by calling `border.toJSON()`. */
        export interface BorderData {
            
            
            
            
        }
        /** An interface describing the data returned by calling `borders.toJSON()`. */
        export interface BordersData {
        }
        /** An interface describing the data returned by calling `margins.toJSON()`. */
        export interface MarginsData {
            
            
            
            
        }
        /** An interface describing the data returned by calling `shapeFill.toJSON()`. */
        export interface ShapeFillData {
            
            
            
        }
        /** An interface describing the data returned by calling `tableCell.toJSON()`. */
        export interface TableCellData {
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `tableCellCollection.toJSON()`. */
        export interface TableCellCollectionData {
            items?: PowerPoint.Interfaces.TableCellData[];
        }
        /** An interface describing the data returned by calling `tableColumn.toJSON()`. */
        export interface TableColumnData {
            
            
        }
        /** An interface describing the data returned by calling `tableColumnCollection.toJSON()`. */
        export interface TableColumnCollectionData {
            items?: PowerPoint.Interfaces.TableColumnData[];
        }
        /** An interface describing the data returned by calling `tableRow.toJSON()`. */
        export interface TableRowData {
            
            
            
        }
        /** An interface describing the data returned by calling `tableRowCollection.toJSON()`. */
        export interface TableRowCollectionData {
            items?: PowerPoint.Interfaces.TableRowData[];
        }
        /** An interface describing the data returned by calling `tableStyleSettings.toJSON()`. */
        export interface TableStyleSettingsData {
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `table.toJSON()`. */
        export interface TableData {
            
            
            
        }
        /** An interface describing the data returned by calling `shapeCollection.toJSON()`. */
        export interface ShapeCollectionData {
            items?: PowerPoint.Interfaces.ShapeData[];
        }
        /** An interface describing the data returned by calling `slideBackgroundGradientFill.toJSON()`. */
        export interface SlideBackgroundGradientFillData {
            
        }
        /** An interface describing the data returned by calling `slideBackgroundPatternFill.toJSON()`. */
        export interface SlideBackgroundPatternFillData {
            
            
            
        }
        /** An interface describing the data returned by calling `slideBackgroundPictureOrTextureFill.toJSON()`. */
        export interface SlideBackgroundPictureOrTextureFillData {
            
        }
        /** An interface describing the data returned by calling `slideBackgroundSolidFill.toJSON()`. */
        export interface SlideBackgroundSolidFillData {
            
            
        }
        /** An interface describing the data returned by calling `slideBackgroundFill.toJSON()`. */
        export interface SlideBackgroundFillData {
            
        }
        /** An interface describing the data returned by calling `slideBackground.toJSON()`. */
        export interface SlideBackgroundData {
            
            
        }
        /** An interface describing the data returned by calling `slideLayoutBackground.toJSON()`. */
        export interface SlideLayoutBackgroundData {
            
            
        }
        /** An interface describing the data returned by calling `slideLayout.toJSON()`. */
        export interface SlideLayoutData {
            
            
            
        }
        /** An interface describing the data returned by calling `slideLayoutCollection.toJSON()`. */
        export interface SlideLayoutCollectionData {
            items?: PowerPoint.Interfaces.SlideLayoutData[];
        }
        /** An interface describing the data returned by calling `slideMasterBackground.toJSON()`. */
        export interface SlideMasterBackgroundData {
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
            /**
             * Gets the unique ID of the slide.
             *
             * @remarks
             * [Api set: PowerPointApi 1.2]
             */
            id?: string;
            
        }
        /** An interface describing the data returned by calling `shapeGroup.toJSON()`. */
        export interface ShapeGroupData {
            
            
        }
        /** An interface describing the data returned by calling `shapeLineFormat.toJSON()`. */
        export interface ShapeLineFormatData {
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `shape.toJSON()`. */
        export interface ShapeData {
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `binding.toJSON()`. */
        export interface BindingData {
            
            
        }
        /** An interface describing the data returned by calling `bindingCollection.toJSON()`. */
        export interface BindingCollectionData {
            items?: PowerPoint.Interfaces.BindingData[];
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
        /** An interface describing the data returned by calling `pageSetup.toJSON()`. */
        export interface PageSetupData {
            
            
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
         * The `Presentation` object is the top-level object with one or more slides that contain the contents of the presentation.
                    To learn more about the PowerPoint object model,
                    see {@link https://learn.microsoft.com/office/dev/add-ins/powerpoint/core-concepts | PowerPoint JavaScript object model}.
         *
         * @remarks
         * [Api set: PowerPointApi 1.0]
         */
        export interface PresentationLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            
            
            /**
             * Returns the title of the presentation.
             *
             * @remarks
             * [Api set: PowerPointApi 1.0]
             */
            title?: boolean;
        }
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        /**
         * Represents a single slide of a presentation.
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         */
        export interface SlideLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            
            
            /**
             * Gets the unique ID of the slide.
             *
             * @remarks
             * [Api set: PowerPointApi 1.2]
             */
            id?: boolean;
            
        }
        
        
        
        
        
        
        
        
        
        /**
         * Represents the collection of slides in the presentation.
         *
         * @remarks
         * [Api set: PowerPointApi 1.2]
         */
        export interface SlideCollectionLoadOptions {
            /**
              Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
             */
            $all?: boolean;
            
            
            
            /**
             * For EACH ITEM in the collection: Gets the unique ID of the slide.
             *
             * @remarks
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
     * Creates and opens a new presentation. Optionally, the presentation can be prepopulated with a Base64-encoded .pptx file.
     *
     * [Api set: PowerPointApi 1.1]
     *
     * @param base64File - Optional. The Base64-encoded .pptx file. The default value is null. The maximum length of the string is 71,680,000 characters.
     */
    export function createPresentation(base64File?: string): Promise<void>;
}


////////////////////////////////////////////////////////////////
///////////////////// End PowerPoint APIs //////////////////////
////////////////////////////////////////////////////////////////