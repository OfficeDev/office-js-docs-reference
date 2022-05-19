import { OfficeExtension } from "../../api-extractor-inputs-office/office"
import { Office as Outlook} from "../../api-extractor-inputs-outlook/outlook"
////////////////////////////////////////////////////////////////
/////////////////////// Begin Excel APIs ///////////////////////
////////////////////////////////////////////////////////////////



export declare namespace Excel {
    
    
    
    
    
    
    
    
    
    
    
    
        grayDownArrow: Icon;
        graySideArrow: Icon;
        grayUpArrow: Icon;
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    export interface Session {
    }
    /**
     * The RequestContext object facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the request context is required to get access to the Excel object model from the add-in.
     */
    export class RequestContext extends OfficeExtension.ClientRequestContext {
        constructor(url?: string | Session);
        readonly workbook: Workbook;
        readonly application: Application;
        
    }
    /**
     * Executes a batch script that performs actions on the Excel object model, using a new RequestContext. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the RequestContext is required to get access to the Excel object model from the add-in.
     */
    export function run<T>(batch: (context: Excel.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the Excel object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     * @param object - A previously-created API object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the RequestContext is required to get access to the Excel object model from the add-in.
     */
    export function run<T>(object: OfficeExtension.ClientObject, batch: (context: Excel.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the Excel object model, using the RequestContext of previously-created API objects.
     * @param objects - An array of previously-created API objects. The array will be validated to make sure that all of the objects share the same context. The batch will use this shared RequestContext, which means that any changes applied to these objects will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the RequestContext is required to get access to the Excel object model from the add-in.
     */
    export function run<T>(objects: OfficeExtension.ClientObject[], batch: (context: Excel.RequestContext) => Promise<T>): Promise<T>;
    /**
    * Executes a batch script that performs actions on the Excel object model, using the RequestContext of a previously-created API object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
    * @param options - The additional options for this Excel.run which specify previous objects, whether to delay the request for cell edit, session info, etc.
    * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the RequestContext is required to get access to the Excel object model from the add-in.
    */
    export function run<T>(options: Excel.RunOptions, batch: (context: Excel.RequestContext) => Promise<T>): Promise<T>;
    /**
     * Executes a batch script that performs actions on the Excel object model, using the RequestContext of a previously-created object. When the promise is resolved, any tracked objects that were automatically allocated during execution will be released.
     *
     * @param context - A previously-created object. The batch will use the same RequestContext as the passed-in object, which means that any changes applied to the object will be picked up by "context.sync()".
     * @param batch - A function that takes in a RequestContext and returns a promise (typically, just the result of "context.sync()"). The context parameter facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, the RequestContext is required to get access to the Excel object model from the add-in.
     */
    export function run<T>(context: OfficeExtension.ClientRequestContext, batch: (context: Excel.RequestContext) => Promise<T>): Promise<T>;
    export function postprocessBindingDescriptor(response: any): any;
    export function getDataCommonPostprocess(response: any, callArgs: any): any;
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    /**
     * Represents the Excel application that manages the workbook.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class Application extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.Application object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ApplicationData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ApplicationData;
    }
    
    /**
     * Workbook is the top level object which contains related workbook objects such as worksheets, tables, and ranges.
                To learn more about the workbook object model, read {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-workbooks | Work with workbooks using the Excel JavaScript API}.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class Workbook extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the Excel application instance that contains this workbook.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly application: Excel.Application;
        /**
         * Represents a collection of bindings that are part of the workbook.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly bindings: Excel.BindingCollection;
        
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.WorkbookProtection object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.WorkbookProtectionData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.WorkbookProtectionData;
    }
    
    /**
     * An Excel worksheet is a grid of cells. It can contain data, tables, charts, etc.
                To learn more about the worksheet object model, read {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-worksheets | Work with worksheets using the Excel JavaScript API}.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class Worksheet extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        
        
        
    }
    
    /**
     * Range represents a set of one or more contiguous cells such as a cell, a row, a column, or a block of cells.
                To learn more about how ranges are used throughout the API, start with {@link https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-core-concepts#ranges | Ranges in the Excel JavaScript API}.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class Range extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.RangeView object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.RangeViewData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.RangeViewData;
    }
    
    /**
     * A collection of all the `NamedItem` objects that are part of the workbook or worksheet, depending on how it was reached.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class NamedItemCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Excel.NamedItem[];
        
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.NamedItemCollectionLoadOptions & Excel.Interfaces.CollectionLoadOptions): Excel.NamedItemCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.NamedItemCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Excel.NamedItemCollection;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `Excel.NamedItemCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.NamedItemCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
        */
        toJSON(): Excel.Interfaces.NamedItemCollectionData;
    }
    /**
     * Represents a defined name for a range of cells or value. Names can be primitive named objects (as seen in the type below), range object, or a reference to a range. This object can be used to obtain range object associated with names.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class NamedItem extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.NamedItem object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.NamedItemData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.NamedItemData;
    }
    
    /**
     * Represents an Office.js binding that is defined in the workbook.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class Binding extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the binding identifier.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly id: string;
        /**
         * Returns the type of the binding. See `Excel.BindingType` for details.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly type: Excel.BindingType | "Range" | "Table" | "Text";
        
        
        /** Gets the loaded child items in this collection. */
        readonly items: Excel.Binding[];
        /**
         * Returns the number of bindings in the collection.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly count: number;
        
        /** Gets the loaded child items in this collection. */
        readonly items: Excel.Table[];
        /**
         * Returns the number of tables in the workbook.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly count: number;
        /**
         * Creates a new table. The range object or source address determines the worksheet under which the table will be added. If the table cannot be added (e.g., because the address is invalid, or the table would overlap with another table), an error will be thrown.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param address - A `Range` object, or a string address or name of the range representing the data source. If the address does not contain a sheet name, the currently-active sheet is used. [Api set: ExcelApi 1.1 / 1.3.  Prior to ExcelApi 1.3, this parameter must be a string. Starting with Excel Api 1.3, this parameter may be a Range object or a string.]
         * @param hasHeaders - A boolean value that indicates whether the data being imported has column labels. If the source does not contain headers (i.e., when this property set to `false`), Excel will automatically generate a header and shift the data down by one row.
         */
        add(address: Range | string, hasHeaders: boolean): Excel.Table;
        
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.TableCollectionLoadOptions & Excel.Interfaces.CollectionLoadOptions): Excel.TableCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.TableCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Excel.TableCollection;
        
        /** Gets the loaded child items in this collection. */
        readonly items: Excel.Table[];
        
        
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.TableColumn object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.TableColumnData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.TableColumnData;
    }
    /**
     * Represents a collection of all the rows that are part of the table.
                
                 Note that unlike ranges or columns, which will adjust if new rows or columns are added before them,
                 a `TableRow` object represents the physical location of the table row, but not the data.
                 That is, if the data is sorted or if new rows are added, a table row will continue
                 to point at the index for which it was created.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class TableRowCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Excel.TableRow[];
        /**
         * Returns the number of rows in the table.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly count: number;
        /**
         * Adds one or more rows to the table. The return object will be the top of the newly added row(s).
                    
                     Note that unlike ranges or columns, which will adjust if new rows or columns are added before them,
                     a `TableRow` object represents the physical location of the table row, but not the data.
                     That is, if the data is sorted or if new rows are added, a table row will continue
                     to point at the index for which it was created.
         *
         * @remarks
         * [Api set: ExcelApi 1.1 for adding a single row; 1.4 allows adding of multiple rows.]
         *
         * @param index - Optional. Specifies the relative position of the new row. If null or -1, the addition happens at the end. Any rows below the inserted row are shifted downwards. Zero-indexed.
         * @param values - Optional. A 2D array of unformatted values of the table row.
         */
        add(index?: number, values?: Array<Array<boolean | string | number>> | boolean | string | number, alwaysInsert?: boolean): Excel.TableRow;
        
        /**
         * Returns the index number of the row within the rows collection of the table. Zero-indexed.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly index: number;
        /**
         * Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.
                    If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        values: any[][];
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.TableRowUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.TableRow): void;
        /**
         * Deletes the row from the table.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        delete(): void;
        /**
         * Returns the range object associated with the entire row.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        getRange(): Excel.Range;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.TableRowLoadOptions): Excel.TableRow;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.TableRow;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.TableRow;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.TableRow object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.TableRowData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.TableRowData;
    }
    
    
    
    
    
    
    
    
    
    /**
     * A format object encapsulating the range's font, fill, borders, alignment, and other properties.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class RangeFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Collection of border objects that apply to the overall range.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly borders: Excel.RangeBorderCollection;
        /**
         * Returns the fill object defined on the overall range.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly fill: Excel.RangeFill;
        /**
         * Returns the font object defined on the overall range.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly font: Excel.RangeFont;
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.FormatProtection object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.FormatProtectionData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.FormatProtectionData;
    }
    /**
     * Represents the background of a range object.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class RangeFill extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * HTML color code representing the color of the background, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange")
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        color: string;
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.RangeFill object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.RangeFillData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.RangeFillData;
    }
    /**
     * Represents the border of an object.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class RangeBorder extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange").
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        color: string;
        /**
         * Constant value that indicates the specific side of the border. See `Excel.BorderIndex` for details.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly sideIndex: Excel.BorderIndex | "EdgeTop" | "EdgeBottom" | "EdgeLeft" | "EdgeRight" | "InsideVertical" | "InsideHorizontal" | "DiagonalDown" | "DiagonalUp";
        /**
         * One of the constants of line style specifying the line style for the border. See `Excel.BorderLineStyle` for details.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        style: Excel.BorderLineStyle | "None" | "Continuous" | "Dash" | "DashDot" | "DashDotDot" | "Dot" | "Double" | "SlantDashDot";
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.RangeBorder object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.RangeBorderData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.RangeBorderData;
    }
    /**
     * Represents the border objects that make up the range border.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class RangeBorderCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Excel.RangeBorder[];
        /**
         * Number of border objects in the collection.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly count: number;
        
        /**
         * Represents the bold status of the font.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        bold: boolean;
        /**
         * HTML color code representation of the text color (e.g., #FF0000 represents Red).
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        color: string;
        /**
         * Specifies the italic status of the font.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        italic: boolean;
        /**
         * Font name (e.g., "Calibri"). The name's length should not be greater than 31 characters.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        name: string;
        /**
         * Font size.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        size: number;
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.RangeFont object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.RangeFontData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.RangeFontData;
    }
    /**
     * A collection of all the chart objects on a worksheet.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Excel.Chart[];
        /**
         * Returns the number of charts in the worksheet.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly count: number;
        /**
         * Creates a new chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param type - Represents the type of a chart. See `Excel.ChartType` for details.
         * @param sourceData - The `Range` object corresponding to the source data.
         * @param seriesBy - Optional. Specifies the way columns or rows are used as data series on the chart. See `Excel.ChartSeriesBy` for details.
         */
        add(type: Excel.ChartType, sourceData: Range, seriesBy?: Excel.ChartSeriesBy): Excel.Chart;
        /**
         * Creates a new chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param typeString - Represents the type of a chart. See `Excel.ChartType` for details.
         * @param sourceData - The `Range` object corresponding to the source data.
         * @param seriesBy - Optional. Specifies the way columns or rows are used as data series on the chart. See `Excel.ChartSeriesBy` for details.
         */
        add(typeString: "Invalid" | "ColumnClustered" | "ColumnStacked" | "ColumnStacked100" | "3DColumnClustered" | "3DColumnStacked" | "3DColumnStacked100" | "BarClustered" | "BarStacked" | "BarStacked100" | "3DBarClustered" | "3DBarStacked" | "3DBarStacked100" | "LineStacked" | "LineStacked100" | "LineMarkers" | "LineMarkersStacked" | "LineMarkersStacked100" | "PieOfPie" | "PieExploded" | "3DPieExploded" | "BarOfPie" | "XYScatterSmooth" | "XYScatterSmoothNoMarkers" | "XYScatterLines" | "XYScatterLinesNoMarkers" | "AreaStacked" | "AreaStacked100" | "3DAreaStacked" | "3DAreaStacked100" | "DoughnutExploded" | "RadarMarkers" | "RadarFilled" | "Surface" | "SurfaceWireframe" | "SurfaceTopView" | "SurfaceTopViewWireframe" | "Bubble" | "Bubble3DEffect" | "StockHLC" | "StockOHLC" | "StockVHLC" | "StockVOHLC" | "CylinderColClustered" | "CylinderColStacked" | "CylinderColStacked100" | "CylinderBarClustered" | "CylinderBarStacked" | "CylinderBarStacked100" | "CylinderCol" | "ConeColClustered" | "ConeColStacked" | "ConeColStacked100" | "ConeBarClustered" | "ConeBarStacked" | "ConeBarStacked100" | "ConeCol" | "PyramidColClustered" | "PyramidColStacked" | "PyramidColStacked100" | "PyramidBarClustered" | "PyramidBarStacked" | "PyramidBarStacked100" | "PyramidCol" | "3DColumn" | "Line" | "3DLine" | "3DPie" | "Pie" | "XYScatter" | "3DArea" | "Area" | "Doughnut" | "Radar" | "Histogram" | "Boxwhisker" | "Pareto" | "RegionMap" | "Treemap" | "Waterfall" | "Sunburst" | "Funnel", sourceData: Range, seriesBy?: "Auto" | "Columns" | "Rows"): Excel.Chart;
        
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartCollectionLoadOptions & Excel.Interfaces.CollectionLoadOptions): Excel.ChartCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartCollection;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Excel.ChartCollection;
        
        /**
         * Represents chart axes.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly axes: Excel.ChartAxes;
        /**
         * Represents the data labels on the chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly dataLabels: Excel.ChartDataLabels;
        /**
         * Encapsulates the format properties for the chart area.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly format: Excel.ChartAreaFormat;
        /**
         * Represents the legend for the chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly legend: Excel.ChartLegend;
        
        
        
        /** Gets the loaded child items in this collection. */
        readonly items: Excel.ChartSeries[];
        /**
         * Returns the number of series in the collection.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly count: number;
        
        
        /**
         * Represents the fill format of a chart series, which includes background formatting information.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly fill: Excel.ChartFill;
        /**
         * Represents line formatting.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly line: Excel.ChartLineFormat;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartSeriesFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartSeriesFormat): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartSeriesFormatLoadOptions): Excel.ChartSeriesFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartSeriesFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartSeriesFormat;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.ChartSeriesFormat object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartSeriesFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartSeriesFormatData;
    }
    /**
     * A collection of all the chart points within a series inside a chart.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartPointsCollection extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /** Gets the loaded child items in this collection. */
        readonly items: Excel.ChartPoint[];
        /**
         * Returns the number of chart points in the series.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly count: number;
        
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.ChartPoint object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartPointData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartPointData;
    }
    /**
     * Represents the formatting object for chart points.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartPointFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.ChartPointFormat object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartPointFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartPointFormatData;
    }
    /**
     * Represents the chart axes.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartAxes extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the category axis in a chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly categoryAxis: Excel.ChartAxis;
        /**
         * Represents the series axis of a 3-D chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly seriesAxis: Excel.ChartAxis;
        /**
         * Represents the value axis in an axis.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly valueAxis: Excel.ChartAxis;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartAxesUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartAxes): void;
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.ChartAxes object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartAxesData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartAxesData;
    }
    /**
     * Represents a single axis in a chart.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartAxis extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the formatting of a chart object, which includes line and font formatting.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly format: Excel.ChartAxisFormat;
        /**
         * Returns an object that represents the major gridlines for the specified axis.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly majorGridlines: Excel.ChartGridlines;
        /**
         * Returns an object that represents the minor gridlines for the specified axis.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly minorGridlines: Excel.ChartGridlines;
        /**
         * Represents the axis title.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly title: Excel.ChartAxisTitle;
        
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.ChartAxisFormat object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartAxisFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartAxisFormatData;
    }
    /**
     * Represents the title of a chart axis.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartAxisTitle extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Specifies the formatting of the chart axis title.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly format: Excel.ChartAxisTitleFormat;
        /**
         * Specifies the axis title.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        text: string;
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.ChartAxisTitle object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartAxisTitleData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartAxisTitleData;
    }
    /**
     * Represents the chart axis title formatting.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartAxisTitleFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.ChartAxisTitleFormat object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartAxisTitleFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartAxisTitleFormatData;
    }
    /**
     * Represents a collection of all the data labels on a chart point.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartDataLabels extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Specifies the format of chart data labels, which includes fill and font formatting.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly format: Excel.ChartDataLabelFormat;
        
        
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.ChartDataLabelFormat object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartDataLabelFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartDataLabelFormatData;
    }
    
    
    
    
    /**
     * Represents major or minor gridlines on a chart axis.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartGridlines extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the formatting of chart gridlines.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly format: Excel.ChartGridlinesFormat;
        /**
         * Specifies if the axis gridlines are visible.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        visible: boolean;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartGridlinesUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartGridlines): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartGridlinesLoadOptions): Excel.ChartGridlines;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartGridlines;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartGridlines;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.ChartGridlines object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartGridlinesData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartGridlinesData;
    }
    /**
     * Encapsulates the format properties for chart gridlines.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartGridlinesFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents chart line formatting.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly line: Excel.ChartLineFormat;
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartGridlinesFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartGridlinesFormat): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartGridlinesFormatLoadOptions): Excel.ChartGridlinesFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartGridlinesFormat;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartGridlinesFormat;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.ChartGridlinesFormat object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartGridlinesFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartGridlinesFormatData;
    }
    /**
     * Represents the legend in a chart.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartLegend extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the formatting of a chart legend, which includes fill and font formatting.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly format: Excel.ChartLegendFormat;
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.ChartLegend object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartLegendData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartLegendData;
    }
    
    
    /**
     * Encapsulates the format properties of a chart legend.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartLegendFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.ChartLegendFormat object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartLegendFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartLegendFormatData;
    }
    
    /**
     * Represents a chart title object of a chart.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartTitle extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the formatting of a chart title, which includes fill and font formatting.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        readonly format: Excel.ChartTitleFormat;
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.ChartTitle object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartTitleData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartTitleData;
    }
    
    /**
     * Provides access to the formatting options for a chart title.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartTitleFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.ChartTitleFormat object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartTitleFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartTitleFormatData;
    }
    /**
     * Represents the fill formatting for a chart element.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartFill extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Clears the fill color of a chart element.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        clear(): void;
        /**
         * Sets the fill formatting of a chart element to a uniform color.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         *
         * @param color - HTML color code representing the color of the background, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").
         */
        setSolidColor(color: string): void;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.ChartFill object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartFillData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): {
            [key: string]: string;
        };
    }
    
    
    
    /**
     * Encapsulates the formatting options for line elements.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartLineFormat extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * HTML color code representing the color of lines in the chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        color: string;
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.ChartLineFormat object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartLineFormatData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartLineFormatData;
    }
    /**
     * This object represents the font attributes (such as font name, font size, and color) for a chart object.
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    export class ChartFont extends OfficeExtension.ClientObject {
        /** The request context associated with the object. This connects the add-in's process to the Office host application's process. */
        context: RequestContext;
        /**
         * Represents the bold status of font.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        bold: boolean;
        /**
         * HTML color code representation of the text color (e.g., #FF0000 represents Red).
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        color: string;
        /**
         * Represents the italic status of the font.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        italic: boolean;
        /**
         * Font name (e.g., "Calibri")
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        name: string;
        /**
         * Size of the font (e.g., 11)
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        size: number;
        /**
         * Type of underline applied to the font. See `Excel.ChartUnderlineStyle` for details.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        underline: Excel.ChartUnderlineStyle | "None" | "Single";
        /**
         * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
         * @param properties - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
         * @param options - Provides an option to suppress errors if the properties object tries to set any read-only properties.
         */
        set(properties: Interfaces.ChartFontUpdateData, options?: OfficeExtension.UpdateOptions): void;
        /** Sets multiple properties on the object at the same time, based on an existing loaded object. */
        set(properties: Excel.ChartFont): void;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param options - Provides options for which properties of the object to load.
         */
        load(options?: Excel.Interfaces.ChartFontLoadOptions): Excel.ChartFont;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNames - A comma-delimited string or an array of strings that specify the properties to load.
         */
        load(propertyNames?: string | string[]): Excel.ChartFont;
        /**
         * Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
         *
         * @param propertyNamesAndPaths - `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
         */
        load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Excel.ChartFont;
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original Excel.ChartFont object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Excel.Interfaces.ChartFontData`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Excel.Interfaces.ChartFontData;
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    /**
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum ChartDataLabelPosition {
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        invalid = "Invalid",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        none = "None",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        center = "Center",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        insideEnd = "InsideEnd",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        insideBase = "InsideBase",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        outsideEnd = "OutsideEnd",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        left = "Left",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        right = "Right",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        top = "Top",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        bottom = "Bottom",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        bestFit = "BestFit",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        callout = "Callout"
    }
    
    
    
    
    
    
    /**
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum ChartLegendPosition {
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        invalid = "Invalid",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        top = "Top",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        bottom = "Bottom",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        left = "Left",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        right = "Right",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        corner = "Corner",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        custom = "Custom"
    }
    
    
    
    
    
    /**
     * Specifies whether the series are by rows or by columns. In Excel on desktop, the "auto" option will inspect the source data shape to automatically guess whether the data is by rows or columns. In Excel on the web, "auto" will simply default to "columns".
     *
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum ChartSeriesBy {
        /**
         * In Excel on desktop, the "auto" option will inspect the source data shape to automatically guess whether the data is by rows or columns. In Excel on the web, "auto" will simply default to "columns".
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        auto = "Auto",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        columns = "Columns",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        rows = "Rows"
    }
    
    
    
    /**
     * @remarks
     * [Api set: ExcelApi 1.1]
     */
    enum ChartType {
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        invalid = "Invalid",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        columnClustered = "ColumnClustered",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        columnStacked = "ColumnStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        columnStacked100 = "ColumnStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DColumnClustered = "3DColumnClustered",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DColumnStacked = "3DColumnStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DColumnStacked100 = "3DColumnStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        barClustered = "BarClustered",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        barStacked = "BarStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        barStacked100 = "BarStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DBarClustered = "3DBarClustered",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DBarStacked = "3DBarStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DBarStacked100 = "3DBarStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        lineStacked = "LineStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        lineStacked100 = "LineStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        lineMarkers = "LineMarkers",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        lineMarkersStacked = "LineMarkersStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        lineMarkersStacked100 = "LineMarkersStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        pieOfPie = "PieOfPie",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        pieExploded = "PieExploded",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DPieExploded = "3DPieExploded",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        barOfPie = "BarOfPie",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        xyscatterSmooth = "XYScatterSmooth",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        xyscatterSmoothNoMarkers = "XYScatterSmoothNoMarkers",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        xyscatterLines = "XYScatterLines",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        xyscatterLinesNoMarkers = "XYScatterLinesNoMarkers",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        areaStacked = "AreaStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        areaStacked100 = "AreaStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DAreaStacked = "3DAreaStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DAreaStacked100 = "3DAreaStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        doughnutExploded = "DoughnutExploded",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        radarMarkers = "RadarMarkers",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        radarFilled = "RadarFilled",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        surface = "Surface",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        surfaceWireframe = "SurfaceWireframe",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        surfaceTopView = "SurfaceTopView",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        surfaceTopViewWireframe = "SurfaceTopViewWireframe",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        bubble = "Bubble",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        bubble3DEffect = "Bubble3DEffect",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        stockHLC = "StockHLC",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        stockOHLC = "StockOHLC",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        stockVHLC = "StockVHLC",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        stockVOHLC = "StockVOHLC",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        cylinderColClustered = "CylinderColClustered",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        cylinderColStacked = "CylinderColStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        cylinderColStacked100 = "CylinderColStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        cylinderBarClustered = "CylinderBarClustered",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        cylinderBarStacked = "CylinderBarStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        cylinderBarStacked100 = "CylinderBarStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        cylinderCol = "CylinderCol",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        coneColClustered = "ConeColClustered",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        coneColStacked = "ConeColStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        coneColStacked100 = "ConeColStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        coneBarClustered = "ConeBarClustered",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        coneBarStacked = "ConeBarStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        coneBarStacked100 = "ConeBarStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        coneCol = "ConeCol",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        pyramidColClustered = "PyramidColClustered",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        pyramidColStacked = "PyramidColStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        pyramidColStacked100 = "PyramidColStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        pyramidBarClustered = "PyramidBarClustered",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        pyramidBarStacked = "PyramidBarStacked",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        pyramidBarStacked100 = "PyramidBarStacked100",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        pyramidCol = "PyramidCol",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DColumn = "3DColumn",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        line = "Line",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DLine = "3DLine",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DPie = "3DPie",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        pie = "Pie",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        xyscatter = "XYScatter",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        _3DArea = "3DArea",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        area = "Area",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        doughnut = "Doughnut",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        radar = "Radar",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        histogram = "Histogram",
         1.7 for Array]
     */
    enum NamedItemType {
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        string = "String",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        integer = "Integer",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        double = "Double",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        boolean = "Boolean",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        range = "Range",
        /**
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        error = "Error",
        
        
        /**
        * Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that is passed to it.)
        * Whereas the original `FunctionResult<T>` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Interfaces.FunctionResultData<T>`) that contains shallow copies of any loaded child properties from the original object.
        */
        toJSON(): Interfaces.FunctionResultData<T>;
    }
    
    enum ErrorCodes {
        accessDenied = "AccessDenied",
        apiNotFound = "ApiNotFound",
        conflict = "Conflict",
        emptyChartSeries = "EmptyChartSeries",
        filteredRangeConflict = "FilteredRangeConflict",
        formulaLengthExceedsLimit = "FormulaLengthExceedsLimit",
        generalException = "GeneralException",
        inactiveWorkbook = "InactiveWorkbook",
        insertDeleteConflict = "InsertDeleteConflict",
        invalidArgument = "InvalidArgument",
        invalidBinding = "InvalidBinding",
        invalidOperation = "InvalidOperation",
        invalidReference = "InvalidReference",
        invalidSelection = "InvalidSelection",
        itemAlreadyExists = "ItemAlreadyExists",
        itemNotFound = "ItemNotFound",
        mergedRangeConflict = "MergedRangeConflict",
        nonBlankCellOffSheet = "NonBlankCellOffSheet",
        notImplemented = "NotImplemented",
        openWorkbookLinksBlocked = "OpenWorkbookLinksBlocked",
        operationCellsExceedLimit = "OperationCellsExceedLimit",
        pivotTableRangeConflict = "PivotTableRangeConflict",
        rangeExceedsLimit = "RangeExceedsLimit",
        refreshWorkbookLinksBlocked = "RefreshWorkbookLinksBlocked",
        requestAborted = "RequestAborted",
        responsePayloadSizeLimitExceeded = "ResponsePayloadSizeLimitExceeded",
        unsupportedFeature = "UnsupportedFeature",
        unsupportedOperation = "UnsupportedOperation",
        unsupportedSheet = "UnsupportedSheet",
        invalidOperationInCellEditMode = "InvalidOperationInCellEditMode"
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
        /** An interface for updating data on the QueryCollection object, for use in `queryCollection.set({ ... })`. */
        export interface QueryCollectionUpdateData {
            items?: Excel.Interfaces.QueryData[];
        }
        /** An interface for updating data on the LinkedWorkbookCollection object, for use in `linkedWorkbookCollection.set({ ... })`. */
        export interface LinkedWorkbookCollectionUpdateData {
            
        }
        /** An interface for updating data on the Application object, for use in `application.set({ ... })`. */
        export interface ApplicationUpdateData {
            
            
            
            
        }
        
        }
        /** An interface for updating data on the SettingCollection object, for use in `settingCollection.set({ ... })`. */
        export interface SettingCollectionUpdateData {
            items?: Excel.Interfaces.SettingData[];
        }
        /** An interface for updating data on the Setting object, for use in `setting.set({ ... })`. */
        export interface SettingUpdateData {
            
        }
        /** An interface for updating data on the NamedItem object, for use in `namedItem.set({ ... })`. */
        export interface NamedItemUpdateData {
            
        }
        /** An interface for updating data on the TableScopedCollection object, for use in `tableScopedCollection.set({ ... })`. */
        export interface TableScopedCollectionUpdateData {
            items?: Excel.Interfaces.TableData[];
        }
        /** An interface for updating data on the Table object, for use in `table.set({ ... })`. */
        export interface TableUpdateData {
            
            
        }
        /** An interface for updating data on the TableColumn object, for use in `tableColumn.set({ ... })`. */
        export interface TableColumnUpdateData {
            /**
             * Specifies the name of the table column.
             *
             * @remarks
             * [Api set: ExcelApi 1.1 for getting the name; 1.4 for setting it.]
             */
            name?: string;
            /**
             * Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.
                        If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            values?: any[][];
        }
        /** An interface for updating data on the TableRowCollection object, for use in `tableRowCollection.set({ ... })`. */
        export interface TableRowCollectionUpdateData {
            items?: Excel.Interfaces.TableRowData[];
        }
        /** An interface for updating data on the TableRow object, for use in `tableRow.set({ ... })`. */
        export interface TableRowUpdateData {
            /**
             * Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.
                        If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            values?: any[][];
        }
        /** An interface for updating data on the DataValidation object, for use in `dataValidation.set({ ... })`. */
        export interface DataValidationUpdateData {
            
            /**
            * Returns the fill object defined on the overall range.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            fill?: Excel.Interfaces.RangeFillUpdateData;
            /**
            * Returns the font object defined on the overall range.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.RangeFontUpdateData;
            
            
            /**
             * One of the constants of line style specifying the line style for the border. See `Excel.BorderLineStyle` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            style?: Excel.BorderLineStyle | "None" | "Continuous" | "Dash" | "DashDot" | "DashDotDot" | "Dot" | "Double" | "SlantDashDot";
            
            items?: Excel.Interfaces.RangeBorderData[];
        }
        /** An interface for updating data on the RangeFont object, for use in `rangeFont.set({ ... })`. */
        export interface RangeFontUpdateData {
            /**
             * Represents the bold status of the font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            bold?: boolean;
            /**
             * HTML color code representation of the text color (e.g., #FF0000 represents Red).
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: string;
            /**
             * Specifies the italic status of the font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            italic?: boolean;
            /**
             * Font name (e.g., "Calibri"). The name's length should not be greater than 31 characters.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: string;
            /**
             * Font size.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            size?: number;
            
        }
        /** An interface for updating data on the Chart object, for use in `chart.set({ ... })`. */
        export interface ChartUpdateData {
            /**
            * Represents chart axes.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            axes?: Excel.Interfaces.ChartAxesUpdateData;
            /**
            * Represents the data labels on the chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            dataLabels?: Excel.Interfaces.ChartDataLabelsUpdateData;
            /**
            * Encapsulates the format properties for the chart area.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartAreaFormatUpdateData;
            /**
            * Represents the legend for the chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            legend?: Excel.Interfaces.ChartLegendUpdateData;
            
            
            /**
            * Represents the font attributes (font name, font size, color, etc.) for the current object.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontUpdateData;
            
        }
        /** An interface for updating data on the ChartSeries object, for use in `chartSeries.set({ ... })`. */
        export interface ChartSeriesUpdateData {
            
        }
        /** An interface for updating data on the ChartPointsCollection object, for use in `chartPointsCollection.set({ ... })`. */
        export interface ChartPointsCollectionUpdateData {
            items?: Excel.Interfaces.ChartPointData[];
        }
        /** An interface for updating data on the ChartPoint object, for use in `chartPoint.set({ ... })`. */
        export interface ChartPointUpdateData {
            
        }
        /** An interface for updating data on the ChartAxes object, for use in `chartAxes.set({ ... })`. */
        export interface ChartAxesUpdateData {
            /**
            * Represents the category axis in a chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            categoryAxis?: Excel.Interfaces.ChartAxisUpdateData;
            /**
            * Represents the series axis of a 3-D chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            seriesAxis?: Excel.Interfaces.ChartAxisUpdateData;
            /**
            * Represents the value axis in an axis.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            valueAxis?: Excel.Interfaces.ChartAxisUpdateData;
        }
        /** An interface for updating data on the ChartAxis object, for use in `chartAxis.set({ ... })`. */
        export interface ChartAxisUpdateData {
            /**
            * Represents the formatting of a chart object, which includes line and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartAxisFormatUpdateData;
            /**
            * Returns an object that represents the major gridlines for the specified axis.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            majorGridlines?: Excel.Interfaces.ChartGridlinesUpdateData;
            /**
            * Returns an object that represents the minor gridlines for the specified axis.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            minorGridlines?: Excel.Interfaces.ChartGridlinesUpdateData;
            /**
            * Represents the axis title.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            title?: Excel.Interfaces.ChartAxisTitleUpdateData;
            
            /**
             * Specifies the axis title.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            text?: string;
            
            /**
            * Specifies the chart axis title's font attributes, such as font name, font size, or color, of the chart axis title object.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontUpdateData;
        }
        /** An interface for updating data on the ChartDataLabels object, for use in `chartDataLabels.set({ ... })`. */
        export interface ChartDataLabelsUpdateData {
            /**
            * Specifies the format of chart data labels, which includes fill and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartDataLabelFormatUpdateData;
            
            /**
             * Specifies if the axis gridlines are visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            visible?: boolean;
        }
        /** An interface for updating data on the ChartGridlinesFormat object, for use in `chartGridlinesFormat.set({ ... })`. */
        export interface ChartGridlinesFormatUpdateData {
            /**
            * Represents chart line formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            line?: Excel.Interfaces.ChartLineFormatUpdateData;
        }
        /** An interface for updating data on the ChartLegend object, for use in `chartLegend.set({ ... })`. */
        export interface ChartLegendUpdateData {
            /**
            * Represents the formatting of a chart legend, which includes fill and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartLegendFormatUpdateData;
            
        }
        /** An interface for updating data on the ChartLegendEntryCollection object, for use in `chartLegendEntryCollection.set({ ... })`. */
        export interface ChartLegendEntryCollectionUpdateData {
            items?: Excel.Interfaces.ChartLegendEntryData[];
        }
        /** An interface for updating data on the ChartLegendFormat object, for use in `chartLegendFormat.set({ ... })`. */
        export interface ChartLegendFormatUpdateData {
            
            
        }
        /** An interface for updating data on the ChartTitleFormat object, for use in `chartTitleFormat.set({ ... })`. */
        export interface ChartTitleFormatUpdateData {
            
            
            
            /**
             * HTML color code representation of the text color (e.g., #FF0000 represents Red).
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: string;
            /**
             * Represents the italic status of the font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            italic?: boolean;
            /**
             * Font name (e.g., "Calibri")
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: string;
            /**
             * Size of the font (e.g., 11)
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            size?: number;
            /**
             * Type of underline applied to the font. See `Excel.ChartUnderlineStyle` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            underline?: Excel.ChartUnderlineStyle | "None" | "Single";
        }
        /** An interface for updating data on the ChartTrendline object, for use in `chartTrendline.set({ ... })`. */
        export interface ChartTrendlineUpdateData {
            
        }
        /** An interface for updating data on the ChartTrendlineLabel object, for use in `chartTrendlineLabel.set({ ... })`. */
        export interface ChartTrendlineLabelUpdateData {
            
            
        }
        /** An interface for updating data on the CustomXmlPartScopedCollection object, for use in `customXmlPartScopedCollection.set({ ... })`. */
        export interface CustomXmlPartScopedCollectionUpdateData {
            items?: Excel.Interfaces.CustomXmlPartData[];
        }
        /** An interface for updating data on the CustomXmlPartCollection object, for use in `customXmlPartCollection.set({ ... })`. */
        export interface CustomXmlPartCollectionUpdateData {
            items?: Excel.Interfaces.CustomXmlPartData[];
        }
        /** An interface for updating data on the PivotTableScopedCollection object, for use in `pivotTableScopedCollection.set({ ... })`. */
        export interface PivotTableScopedCollectionUpdateData {
            items?: Excel.Interfaces.PivotTableData[];
        }
        /** An interface for updating data on the PivotTableCollection object, for use in `pivotTableCollection.set({ ... })`. */
        export interface PivotTableCollectionUpdateData {
            items?: Excel.Interfaces.PivotTableData[];
        }
        /** An interface for updating data on the PivotTable object, for use in `pivotTable.set({ ... })`. */
        export interface PivotTableUpdateData {
            
        }
        /** An interface for updating data on the RowColumnPivotHierarchyCollection object, for use in `rowColumnPivotHierarchyCollection.set({ ... })`. */
        export interface RowColumnPivotHierarchyCollectionUpdateData {
            items?: Excel.Interfaces.RowColumnPivotHierarchyData[];
        }
        /** An interface for updating data on the RowColumnPivotHierarchy object, for use in `rowColumnPivotHierarchy.set({ ... })`. */
        export interface RowColumnPivotHierarchyUpdateData {
            
        }
        /** An interface for updating data on the FilterPivotHierarchy object, for use in `filterPivotHierarchy.set({ ... })`. */
        export interface FilterPivotHierarchyUpdateData {
            
        }
        /** An interface for updating data on the DataPivotHierarchy object, for use in `dataPivotHierarchy.set({ ... })`. */
        export interface DataPivotHierarchyUpdateData {
            
        }
        /** An interface for updating data on the PivotField object, for use in `pivotField.set({ ... })`. */
        export interface PivotFieldUpdateData {
            
        }
        /** An interface for updating data on the PivotItem object, for use in `pivotItem.set({ ... })`. */
        export interface PivotItemUpdateData {
            
            
        }
        /** An interface for updating data on the CustomPropertyCollection object, for use in `customPropertyCollection.set({ ... })`. */
        export interface CustomPropertyCollectionUpdateData {
            items?: Excel.Interfaces.CustomPropertyData[];
        }
        /** An interface for updating data on the ConditionalFormatCollection object, for use in `conditionalFormatCollection.set({ ... })`. */
        export interface ConditionalFormatCollectionUpdateData {
            items?: Excel.Interfaces.ConditionalFormatData[];
        }
        /** An interface for updating data on the ConditionalFormat object, for use in `conditionalFormat.set({ ... })`. */
        export interface ConditionalFormatUpdateData {
            
            
            
        }
        /** An interface for updating data on the ConditionalDataBarPositiveFormat object, for use in `conditionalDataBarPositiveFormat.set({ ... })`. */
        export interface ConditionalDataBarPositiveFormatUpdateData {
            
            
            
            
            
        }
        /** An interface for updating data on the TopBottomConditionalFormat object, for use in `topBottomConditionalFormat.set({ ... })`. */
        export interface TopBottomConditionalFormatUpdateData {
            
            
            
            
            
            
        }
        /** An interface for updating data on the ConditionalRangeBorder object, for use in `conditionalRangeBorder.set({ ... })`. */
        export interface ConditionalRangeBorderUpdateData {
            
            
        }
        /** An interface for updating data on the RangeAreasCollection object, for use in `rangeAreasCollection.set({ ... })`. */
        export interface RangeAreasCollectionUpdateData {
            items?: Excel.Interfaces.RangeAreasData[];
        }
        /** An interface for updating data on the CommentCollection object, for use in `commentCollection.set({ ... })`. */
        export interface CommentCollectionUpdateData {
            items?: Excel.Interfaces.CommentData[];
        }
        /** An interface for updating data on the Comment object, for use in `comment.set({ ... })`. */
        export interface CommentUpdateData {
            
        }
        /** An interface for updating data on the ShapeCollection object, for use in `shapeCollection.set({ ... })`. */
        export interface ShapeCollectionUpdateData {
            items?: Excel.Interfaces.ShapeData[];
        }
        /** An interface for updating data on the Shape object, for use in `shape.set({ ... })`. */
        export interface ShapeUpdateData {
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `application.toJSON()`. */
        export interface ApplicationData {
            
            
            
        }
        /** An interface describing the data returned by calling `workbookCreated.toJSON()`. */
        export interface WorkbookCreatedData {
        }
        /** An interface describing the data returned by calling `worksheet.toJSON()`. */
        export interface WorksheetData {
            
            
        }
        /** An interface describing the data returned by calling `settingCollection.toJSON()`. */
        export interface SettingCollectionData {
            items?: Excel.Interfaces.SettingData[];
        }
        /** An interface describing the data returned by calling `setting.toJSON()`. */
        export interface SettingData {
            
        }
        /** An interface describing the data returned by calling `namedItem.toJSON()`. */
        export interface NamedItemData {
            
            
            /**
             * Returns the type of the binding. See `Excel.BindingType` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            type?: Excel.BindingType | "Range" | "Table" | "Text";
        }
        /** An interface describing the data returned by calling `bindingCollection.toJSON()`. */
        export interface BindingCollectionData {
            items?: Excel.Interfaces.BindingData[];
        }
        /** An interface describing the data returned by calling `tableCollection.toJSON()`. */
        export interface TableCollectionData {
            items?: Excel.Interfaces.TableData[];
        }
        /** An interface describing the data returned by calling `tableScopedCollection.toJSON()`. */
        export interface TableScopedCollectionData {
            items?: Excel.Interfaces.TableData[];
        }
        /** An interface describing the data returned by calling `table.toJSON()`. */
        export interface TableData {
            
            
        }
        /** An interface describing the data returned by calling `tableColumn.toJSON()`. */
        export interface TableColumnData {
            
        }
        /** An interface describing the data returned by calling `tableRow.toJSON()`. */
        export interface TableRowData {
            /**
             * Returns the index number of the row within the rows collection of the table. Zero-indexed.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            index?: number;
            /**
             * Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.
                        If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            values?: any[][];
        }
        /** An interface describing the data returned by calling `dataValidation.toJSON()`. */
        export interface DataValidationData {
            
            /**
            * Returns the font object defined on the overall range.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.RangeFontData;
            
            
            /**
             * Constant value that indicates the specific side of the border. See `Excel.BorderIndex` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            sideIndex?: Excel.BorderIndex | "EdgeTop" | "EdgeBottom" | "EdgeLeft" | "EdgeRight" | "InsideVertical" | "InsideHorizontal" | "DiagonalDown" | "DiagonalUp";
            /**
             * One of the constants of line style specifying the line style for the border. See `Excel.BorderLineStyle` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            style?: Excel.BorderLineStyle | "None" | "Continuous" | "Dash" | "DashDot" | "DashDotDot" | "Dot" | "Double" | "SlantDashDot";
            
        }
        /** An interface describing the data returned by calling `rangeFont.toJSON()`. */
        export interface RangeFontData {
            /**
             * Represents the bold status of the font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            bold?: boolean;
            /**
             * HTML color code representation of the text color (e.g., #FF0000 represents Red).
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: string;
            /**
             * Specifies the italic status of the font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            italic?: boolean;
            /**
             * Font name (e.g., "Calibri"). The name's length should not be greater than 31 characters.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: string;
            /**
             * Font size.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            size?: number;
            
        }
        /** An interface describing the data returned by calling `chart.toJSON()`. */
        export interface ChartData {
            /**
            * Represents chart axes.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            axes?: Excel.Interfaces.ChartAxesData;
            /**
            * Represents the data labels on the chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            dataLabels?: Excel.Interfaces.ChartDataLabelsData;
            /**
            * Encapsulates the format properties for the chart area.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartAreaFormatData;
            /**
            * Represents the legend for the chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            legend?: Excel.Interfaces.ChartLegendData;
            
            
            /**
            * Represents the font attributes (font name, font size, color, etc.) for the current object.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontData;
            
        }
        /** An interface describing the data returned by calling `chartSeries.toJSON()`. */
        export interface ChartSeriesData {
            
        }
        /** An interface describing the data returned by calling `chartPointsCollection.toJSON()`. */
        export interface ChartPointsCollectionData {
            items?: Excel.Interfaces.ChartPointData[];
        }
        /** An interface describing the data returned by calling `chartPoint.toJSON()`. */
        export interface ChartPointData {
            
        }
        /** An interface describing the data returned by calling `chartAxes.toJSON()`. */
        export interface ChartAxesData {
            /**
            * Represents the category axis in a chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            categoryAxis?: Excel.Interfaces.ChartAxisData;
            /**
            * Represents the series axis of a 3-D chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            seriesAxis?: Excel.Interfaces.ChartAxisData;
            /**
            * Represents the value axis in an axis.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            valueAxis?: Excel.Interfaces.ChartAxisData;
        }
        /** An interface describing the data returned by calling `chartAxis.toJSON()`. */
        export interface ChartAxisData {
            /**
            * Represents the formatting of a chart object, which includes line and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartAxisFormatData;
            /**
            * Returns an object that represents the major gridlines for the specified axis.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            majorGridlines?: Excel.Interfaces.ChartGridlinesData;
            /**
            * Returns an object that represents the minor gridlines for the specified axis.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            minorGridlines?: Excel.Interfaces.ChartGridlinesData;
            /**
            * Represents the axis title.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            title?: Excel.Interfaces.ChartAxisTitleData;
            
            /**
             * Specifies the axis title.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            text?: string;
            
            /**
            * Specifies the chart axis title's font attributes, such as font name, font size, or color, of the chart axis title object.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.ChartFontData;
        }
        /** An interface describing the data returned by calling `chartDataLabels.toJSON()`. */
        export interface ChartDataLabelsData {
            /**
            * Specifies the format of chart data labels, which includes fill and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartDataLabelFormatData;
            
            /**
             * Specifies if the axis gridlines are visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            visible?: boolean;
        }
        /** An interface describing the data returned by calling `chartGridlinesFormat.toJSON()`. */
        export interface ChartGridlinesFormatData {
            /**
            * Represents chart line formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            line?: Excel.Interfaces.ChartLineFormatData;
        }
        /** An interface describing the data returned by calling `chartLegend.toJSON()`. */
        export interface ChartLegendData {
            /**
            * Represents the formatting of a chart legend, which includes fill and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartLegendFormatData;
            
            
        }
        /** An interface describing the data returned by calling `chartTitleFormat.toJSON()`. */
        export interface ChartTitleFormatData {
            
            
            
            /**
             * HTML color code representation of the text color (e.g., #FF0000 represents Red).
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: string;
            /**
             * Represents the italic status of the font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            italic?: boolean;
            /**
             * Font name (e.g., "Calibri")
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: string;
            /**
             * Size of the font (e.g., 11)
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            size?: number;
            /**
             * Type of underline applied to the font. See `Excel.ChartUnderlineStyle` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            underline?: Excel.ChartUnderlineStyle | "None" | "Single";
        }
        /** An interface describing the data returned by calling `chartTrendline.toJSON()`. */
        export interface ChartTrendlineData {
            
        }
        /** An interface describing the data returned by calling `chartTrendlineLabel.toJSON()`. */
        export interface ChartTrendlineLabelData {
            
            
        }
        /** An interface describing the data returned by calling `tableSort.toJSON()`. */
        export interface TableSortData {
            
        }
        /** An interface describing the data returned by calling `autoFilter.toJSON()`. */
        export interface AutoFilterData {
            
            
        }
        /** An interface describing the data returned by calling `pivotTableCollection.toJSON()`. */
        export interface PivotTableCollectionData {
            items?: Excel.Interfaces.PivotTableData[];
        }
        /** An interface describing the data returned by calling `pivotTable.toJSON()`. */
        export interface PivotTableData {
            
        }
        /** An interface describing the data returned by calling `rowColumnPivotHierarchy.toJSON()`. */
        export interface RowColumnPivotHierarchyData {
            
        }
        /** An interface describing the data returned by calling `filterPivotHierarchy.toJSON()`. */
        export interface FilterPivotHierarchyData {
            
        }
        /** An interface describing the data returned by calling `dataPivotHierarchy.toJSON()`. */
        export interface DataPivotHierarchyData {
            
        }
        /** An interface describing the data returned by calling `pivotField.toJSON()`. */
        export interface PivotFieldData {
            
        }
        /** An interface describing the data returned by calling `pivotItem.toJSON()`. */
        export interface PivotItemData {
            
            
            
        }
        /** An interface describing the data returned by calling `conditionalFormatCollection.toJSON()`. */
        export interface ConditionalFormatCollectionData {
            items?: Excel.Interfaces.ConditionalFormatData[];
        }
        /** An interface describing the data returned by calling `conditionalFormat.toJSON()`. */
        export interface ConditionalFormatData {
            
            
            
        }
        /** An interface describing the data returned by calling `conditionalDataBarPositiveFormat.toJSON()`. */
        export interface ConditionalDataBarPositiveFormatData {
            
            
            
            
            
            
            
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `conditionalRangeBorder.toJSON()`. */
        export interface ConditionalRangeBorderData {
            
        }
        /** An interface describing the data returned by calling `style.toJSON()`. */
        export interface StyleData {
            
        }
        /** An interface describing the data returned by calling `tableStyle.toJSON()`. */
        export interface TableStyleData {
            
        }
        /** An interface describing the data returned by calling `pivotTableStyle.toJSON()`. */
        export interface PivotTableStyleData {
            
        }
        /** An interface describing the data returned by calling `slicerStyle.toJSON()`. */
        export interface SlicerStyleData {
            
        }
        /** An interface describing the data returned by calling `timelineStyle.toJSON()`. */
        export interface TimelineStyleData {
            
            
        }
        /** An interface describing the data returned by calling `headerFooter.toJSON()`. */
        export interface HeaderFooterData {
            
            
            
        }
        /** An interface describing the data returned by calling `rangeCollection.toJSON()`. */
        export interface RangeCollectionData {
            items?: Excel.Interfaces.RangeData[];
        }
        /** An interface describing the data returned by calling `rangeAreasCollection.toJSON()`. */
        export interface RangeAreasCollectionData {
            items?: Excel.Interfaces.RangeAreasData[];
        }
        /** An interface describing the data returned by calling `commentCollection.toJSON()`. */
        export interface CommentCollectionData {
            items?: Excel.Interfaces.CommentData[];
        }
        /** An interface describing the data returned by calling `comment.toJSON()`. */
        export interface CommentData {
            
            
            
            
            
        }
        /** An interface describing the data returned by calling `line.toJSON()`. */
        export interface LineData {
            
            
            
            
            
            
            
            
            
            /**
             * For EACH ITEM in the collection: Represents the binding identifier.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            id?: boolean;
            /**
             * For EACH ITEM in the collection: Returns the type of the binding. See `Excel.BindingType` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            type?: boolean;
        }
        /**
         * Represents a collection of all the tables that are part of the workbook or worksheet, depending on how it was reached.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface TableCollectionLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            
            
            
            
            
            
            /**
             * For EACH ITEM in the collection: Returns the index number of the row within the rows collection of the table. Zero-indexed.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            index?: boolean;
            /**
             * For EACH ITEM in the collection: Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.
                        If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            values?: boolean;
        }
        /**
         * Represents a row in a table.
                    
                     Note that unlike ranges or columns, which will adjust if new rows or columns are added before them,
                     a `TableRow` object represents the physical location of the table row, but not the data.
                     That is, if the data is sorted or if new rows are added, a table row will continue
                     to point at the index for which it was created.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface TableRowLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
             * Returns the index number of the row within the rows collection of the table. Zero-indexed.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            index?: boolean;
            /**
             * Represents the raw values of the specified range. The data returned could be a string, number, or boolean. Cells that contain an error will return the error string.
                        If the returned value starts with a plus ("+"), minus ("-"), or equal sign ("="), Excel interprets this value as a formula.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            values?: boolean;
        }
        
        
        /**
         * A format object encapsulating the range's font, fill, borders, alignment, and other properties.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface RangeFormatLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Collection of border objects that apply to the overall range.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            borders?: Excel.Interfaces.RangeBorderCollectionLoadOptions;
            /**
            * Returns the fill object defined on the overall range.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            fill?: Excel.Interfaces.RangeFillLoadOptions;
            /**
            * Returns the font object defined on the overall range.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            font?: Excel.Interfaces.RangeFontLoadOptions;
            
            /**
             * HTML color code representing the color of the background, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange")
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: boolean;
            
            /**
             * HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange").
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: boolean;
            /**
             * Constant value that indicates the specific side of the border. See `Excel.BorderIndex` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            sideIndex?: boolean;
            /**
             * One of the constants of line style specifying the line style for the border. See `Excel.BorderLineStyle` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            style?: boolean;
            
            /**
             * For EACH ITEM in the collection: HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500"), or as a named HTML color (e.g., "orange").
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: boolean;
            /**
             * For EACH ITEM in the collection: Constant value that indicates the specific side of the border. See `Excel.BorderIndex` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            sideIndex?: boolean;
            /**
             * For EACH ITEM in the collection: One of the constants of line style specifying the line style for the border. See `Excel.BorderLineStyle` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            style?: boolean;
            
            /**
             * Represents the bold status of the font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            bold?: boolean;
            /**
             * HTML color code representation of the text color (e.g., #FF0000 represents Red).
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: boolean;
            /**
             * Specifies the italic status of the font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            italic?: boolean;
            /**
             * Font name (e.g., "Calibri"). The name's length should not be greater than 31 characters.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: boolean;
            /**
             * Font size.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            size?: boolean;
            
            /**
            * For EACH ITEM in the collection: Represents chart axes.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            axes?: Excel.Interfaces.ChartAxesLoadOptions;
            /**
            * For EACH ITEM in the collection: Represents the data labels on the chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            dataLabels?: Excel.Interfaces.ChartDataLabelsLoadOptions;
            /**
            * For EACH ITEM in the collection: Encapsulates the format properties for the chart area.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartAreaFormatLoadOptions;
            /**
            * For EACH ITEM in the collection: Represents the legend for the chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            legend?: Excel.Interfaces.ChartLegendLoadOptions;
            
            /**
            * Represents chart axes.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            axes?: Excel.Interfaces.ChartAxesLoadOptions;
            /**
            * Represents the data labels on the chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            dataLabels?: Excel.Interfaces.ChartDataLabelsLoadOptions;
            /**
            * Encapsulates the format properties for the chart area.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartAreaFormatLoadOptions;
            /**
            * Represents the legend for the chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            legend?: Excel.Interfaces.ChartLegendLoadOptions;
            
            
            
            
            
            
            /**
            * Represents the category axis in a chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            categoryAxis?: Excel.Interfaces.ChartAxisLoadOptions;
            /**
            * Represents the series axis of a 3-D chart.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            seriesAxis?: Excel.Interfaces.ChartAxisLoadOptions;
            /**
            * Represents the value axis in an axis.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            valueAxis?: Excel.Interfaces.ChartAxisLoadOptions;
        }
        /**
         * Represents a single axis in a chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartAxisLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Represents the formatting of a chart object, which includes line and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartAxisFormatLoadOptions;
            /**
            * Returns an object that represents the major gridlines for the specified axis.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            majorGridlines?: Excel.Interfaces.ChartGridlinesLoadOptions;
            /**
            * Returns an object that represents the minor gridlines for the specified axis.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            minorGridlines?: Excel.Interfaces.ChartGridlinesLoadOptions;
            /**
            * Represents the axis title.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            title?: Excel.Interfaces.ChartAxisTitleLoadOptions;
            
            /**
            * Specifies the formatting of the chart axis title.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartAxisTitleFormatLoadOptions;
            /**
             * Specifies the axis title.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            text?: boolean;
            
            
            /**
            * Specifies the format of chart data labels, which includes fill and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartDataLabelFormatLoadOptions;
            
            /**
            * Represents the formatting of chart gridlines.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartGridlinesFormatLoadOptions;
            /**
             * Specifies if the axis gridlines are visible.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            visible?: boolean;
        }
        /**
         * Encapsulates the format properties for chart gridlines.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartGridlinesFormatLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Represents chart line formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            line?: Excel.Interfaces.ChartLineFormatLoadOptions;
        }
        /**
         * Represents the legend in a chart.
         *
         * @remarks
         * [Api set: ExcelApi 1.1]
         */
        export interface ChartLegendLoadOptions {
            /**
              Specifying `$all` for the LoadOptions loads all the scalar properties (e.g.: `Range.address`) but not the navigational properties (e.g.: `Range.format.fill.color`).
             */
            $all?: boolean;
            /**
            * Represents the formatting of a chart legend, which includes fill and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartLegendFormatLoadOptions;
            
            
            
            
            /**
            * Represents the formatting of a chart title, which includes fill and font formatting.
            *
            * @remarks
            * [Api set: ExcelApi 1.1]
            */
            format?: Excel.Interfaces.ChartTitleFormatLoadOptions;
            
            
            
            
            /**
             * HTML color code representing the color of lines in the chart.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: boolean;
            
            /**
             * Represents the bold status of font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            bold?: boolean;
            /**
             * HTML color code representation of the text color (e.g., #FF0000 represents Red).
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            color?: boolean;
            /**
             * Represents the italic status of the font.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            italic?: boolean;
            /**
             * Font name (e.g., "Calibri")
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            name?: boolean;
            /**
             * Size of the font (e.g., 11)
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            size?: boolean;
            /**
             * Type of underline applied to the font. See `Excel.ChartUnderlineStyle` for details.
             *
             * @remarks
             * [Api set: ExcelApi 1.1]
             */
            underline?: boolean;
        }
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
    }
}


////////////////////////////////////////////////////////////////
//////////////////////// End Excel APIs ////////////////////////
////////////////////////////////////////////////////////////////