// Type definitions for Office.js
// Project: http://dev.office.com
// Definitions by: OfficeDev <https://github.com/OfficeDev>, Lance Austin <https://github.com/LanceEA>, Michael Zlatkovsky <https://github.com/Zlatkovsky>, Kim Brandl <https://github.com/kbrandl>, Ricky Kirkham <https://github.com/Rick-Kirkham>
// Definitions: https://github.com/DefinitelyTyped/DefinitelyTyped

/*
office-js
Copyright (c) Microsoft Corporation
*/

export declare namespace Office {
    export var context: Context;
    /**
     * This method is called after the Office API was loaded.
     * @param reason - Indicates how the app was initialized
     */
    export function initialize(reason: InitializationReason): void;
    /**
     * Indicates if the large namespace for objects will be used or not.
     * @param useShortNamespace  - Indicates if 'true' that the short namespace will be used
     */
    export function useShortNamespace(useShortNamespace: boolean): void;
    // Enumerations
    export enum AsyncResultStatus {
        /**
         * Operation succeeded
         */
        Succeeded,
        /**
         * Operation failed, check error object
         */
        Failed
    }
    export enum InitializationReason {
        /**
         * Indicates the app was just inserted in the document
         */
        Inserted,
        /**
         * Indicates if the extension already existed in the document
         */
        DocumentOpened
    }
    export enum HostType {
        /**
         * Host is Word
         */
        Word,
        /**
         * Host is Excel
         */
        Excel,
        /**
         * Host is PowerPoint
         */
        PowerPoint,
        /**
         * Host is Outlook
         */
        Outlook,
        /**
         * Host is OneNote
         */
        OneNote,
        /**
         * Host is Project
         */
        Project,
        /**
         * Host is Access
         */
        Access
    }
    export enum PlatformType {
        /**
         * Platform is PC
         */
        PC,
        /**
         * Platform is Web
         */
        OfficeOnline,
        /**
         * Platform is Mac
         */
        Mac,
        /**
         * Platform is iOS
         */
        iOS,
        /**
         * Platform is Android
         */
        Android,
        /**
         * Platform is Winrt
         */
        Universal
    }
    // Objects
    export interface AsyncResult {
        asyncContext: any;
        status: AsyncResultStatus;
        error: Error;
        value: any;
    }
    export interface Context {
        auth: Auth;
        contentLanguage: string;
        displayLanguage: string;
        license: string;
        officeTheme: OfficeTheme;
        touchEnabled: boolean;
        ui: UI;
        host: HostType;
        platform: PlatformType;
        diagnostics: {
            host: HostType;
            platform: PlatformType;
            version: string;
        };
        requirements: {
            /**
             * Check if the specified requirement set is supported by the host Office application.
             * @param name - Set name. e.g.: "MatrixBindings".
             * @param minVersion - The minimum required version.
             */
            isSetSupported(name: string, minVersion?: number): boolean;
        }
    }
    /**
     * Provides specific information about an error that occurred during an asynchronous data operation.
     */
    export interface Error {
        /**
         * Gets the numeric code of the error.
         */
        code: number;
        /**
         * Gets the name of the error.
         */
        message: string;
        /**
         * Gets a detailed description of the error.
         */
        name: string;
    }
    export interface UI {
        /**
        * Displays a dialog to show or collect information from the user or to facilitate Web navigation.
        * @param startAddress - Accepts the initial HTTPS Url that opens in the dialog.
        */
        displayDialogAsync(startAddress: string): void;
        /**
        * Displays a dialog to show or collect information from the user or to facilitate Web navigation.
        * @param startAddress - Accepts the initial HTTPS Url that opens in the dialog.
        * @param options - Optional. Accepts a DialogOptions object to define dialog behaviors.
        */
        displayDialogAsync(startAddress: string, options: DialogOptions): void;
        /**
        * Displays a dialog to show or collect information from the user or to facilitate Web navigation.
        * @param startAddress - Accepts the initial HTTPS Url that opens in the dialog.
        * @param callback - Optional. Accepts a callback method to handle the dialog creation attempt.
        */
        displayDialogAsync(startAddress: string, callback: (result: AsyncResult) => void): void;
        /**
        * Displays a dialog to show or collect information from the user or to facilitate Web navigation.
        * @param startAddress - Accepts the initial HTTPS Url that opens in the dialog.
        * @param options - Optional. Accepts a DialogOptions object to define dialog behaviors.
        * @param callback - Optional. Accepts a callback method to handle the dialog creation attempt.
        */
        displayDialogAsync(startAddress: string, options: DialogOptions, callback: (result: AsyncResult) => void): void;

        /**
         * Synchronously delivers a message from the dialog to its parent add-in.
         * @param messageObject - Accepts a message from the dialog to deliver to the add-in.
         */
        messageParent(messageObject: any): void;
        /**
         * Closes the UI container where the JavaScript is executing.
         * 
         * Supported hosts: Outlook - Minimum requirement set: Mailbox 1.5
         * 
         * The behavior of this method is specified by the following:
         * 
         * Called from a UI-less command button: No effect. Any dialog opened by displayDialogAsync will remain open.
         * 
         * Called from a taskpane: The taskpane will close. Any dialog opened by displayDialogAsync will also close. If the taskpane supports pinning and was pinned by the user, it will be un-pinned.
         * 
         * Called from a module extension: No effect.
         */
        closeContainer(): void;
    }
    export interface DialogOptions {
        /**
         * Optional. Defines the width of the dialog as a percentage of the current display. Defaults to 99%. 250px minimum.
         */
        height?: number,
        /**
         * Optional. Defines the height of the dialog as a percentage of the current display. Defaults to 99%. 150px minimum.
         */
        width?: number,
        /**
         * Optional. Determines whether the dialog box should be displayed within an IFrame. This setting is only applicable in Office Online clients, and is ignored on other platforms.
         */
        displayInIframe?: boolean
    }
    export interface Auth {
        /**
        * Obtains an access token from AAD V 2.0 endpoint to grant the Office host application access to the add-in's web application.
        * @param callback - Optional. Accepts a callback method to handle the token acquisition attempt. If AsyncResult.status is "succeeded", then AsyncResult.value is the raw AAD v. 2.0-formatted access token.
        */
        getAccessTokenAsync(callback: (result: AsyncResult) => void): void;
        /**
        * Obtains an access token from AAD V 2.0 endpoint to grant the Office host application access to the add-in's web application.
        * @param options - Optional. Accepts an AuthOptions object to define sign-on behaviors.
        * @param callback - Optional. Accepts a callback method to handle the token acquisition attempt. If AsyncResult.status is "succeeded", then AsyncResult.value is the raw AAD v. 2.0-formatted access token.
        */
        getAccessTokenAsync(options: AuthOptions, callback: (result: AsyncResult) => void): void;

    }
    export interface AuthOptions {
        /**
         * Optional. Causes Office to display the add-in consent experience. Useful if the add-in's Azure permissions have changed or if the user's consent has been revoked.
         */
        forceConsent?: boolean,
        /**
         * Optional. Prompts the user to add (or to switch if already added) his or her Office account.
         */
        forceAddAccount?: boolean,
        /**
         * Optional. Causes Office to prompt the user to provide the additional factor when the tenancy being targeted by Microsoft Graph requires multifactor authentication. The string value identifies the type of additional factor that is required. In most cases, you won't know at development time whether the user's tenant requires an additional factor or what the string should be. So this option would be used in a "second try" call of getAccessTokenAsync after Microsoft Graph has sent an error requesting the additional factor and containing the string that should be used with the authChallenge option.
         */
        authChallenge?: string
        /**
         * Optional. A user-defined item of any type that is returned in the AsyncResult object without being altered.
         */
        asyncContext?: any
    }
    export interface OfficeTheme {
        bodyBackgroundColor: string;
        bodyForegroundColor: string;
        controlBackgroundColor: string;
        controlForegroundColor: string;
    }
    /**
     * Dialog object returned as part of the displayDialogAsync callback. The object exposes methods for registering event handlers and closing the dialog
     */
    export interface DialogHandler {
        /**
         * When called from an active add-in dialog, asynchronously closes the dialog.
         */
        close(): void;
        /**
         * Adds an event handler for DialogMessageReceived or DialogEventReceived
         */
        addEventHandler(eventType: Office.EventType, handler: Function): void;

    }
}

export declare namespace Office {
    /**
     * Returns a promise of an object described in the expression. Callback is invoked only if method fails.
     * @param expression - The object to be retrieved. Example "bindings#BindingName", retrieves a binding promise for a binding named 'BindingName'
     * @param callback - The optional callback method
     */
    export function select(expression: string, callback?: (result: AsyncResult) => void): Binding;
    // Enumerations
    export enum ActiveView {
        Read,
        Edit
    }
    export enum BindingType {
        /**
         * Text based Binding
         */
        Text,
        /**
         * Matrix based Binding
         */
        Matrix,
        /**
         * Table based Binding
         */
        Table
    }
    export enum CoercionType {
        /**
         * Coerce as Text
         */
        Text,
        /**
         * Coerce as Matrix
         */
        Matrix,
        /**
         * Coerce as Table
         */
        Table,
        /**
         * Coerce as HTML
         */
        Html,
        /**
         * Coerce as Office Open XML
         */
        Ooxml,
        /**
         * Coerce as JSON object containing an array of the ids, titles, and indexes of the selected slides.
         */
        SlideRange,
        /**
        * Coerce as Image
        */
        Image
    }
    export enum DocumentMode {
        /**
         * Document in Read Only Mode
         */
        ReadOnly,
        /**
         * Document in Read/Write Mode
         */
        ReadWrite
    }
    export enum EventType {
        /**
         * Occurs when the user changes the current view of the document.
         */
        ActiveViewChanged,
        /**
         * Triggers when a binding level data change happens
         */
        BindingDataChanged,
        /**
         *  Triggers when a binding level selection happens
         */
        BindingSelectionChanged,
        /**
         * Triggers when Dialog sends a message via MessageParent.
         */
        DialogMessageReceived,
        /**
         * Triggers when Dialog has a event, such as dialog closed, dialog navigation failed.
         */
        DialogEventReceived,
        /**
         * Triggers when a document level selection happens
         */
        DocumentSelectionChanged,
        /**
         * Triggers when the active item changes
         */
        ItemChanged,
        /**
         * Triggers when a customXmlPart node was deleted
         */
        NodeDeleted,
        /**
         * Triggers when a customXmlPart node was inserted
         */
        NodeInserted,
        /**
         * Triggers when a customXmlPart node was replaced
         */
        NodeReplaced,
        /**
         * Triggers when settings change in a co-Auth session.
         */
        SettingsChanged,
        /**
         * Triggers when a Task selection happens in Project.
         */
        TaskSelectionChanged,
        /**
         *  Triggers when a Resource selection happens in Project.
         */
        ResourceSelectionChanged,
        /**
         * Triggers when a View selection happens in Project.
         */
        ViewSelectionChanged
    }
    export enum FileType {
        /**
         * Returns the file as plain text
         */
        Text,
        /**
         * Returns the file as a byte array
         */
        Compressed,
        /**
         * Returns the file in PDF format as a byte array
         */
        Pdf
    }
    export enum FilterType {
        /**
         * Returns all items
         */
        All,
        /**
         * Returns only visible items
         */
        OnlyVisible
    }
    export enum GoToType {
        /**
         * Goes to a binding object using the specified binding id.
         */
        Binding,
        /**
         * Goes to a named item using that item's name.
         * In Excel, you can use any structured reference for a named range or table: "Worksheet2!Table1"
         */
        NamedItem,
        /**
         * Goes to a slide using the specified id.
         */
        Slide,
        /**
         * Goes to the specified index by slide number or enum Office.Index
         */
        Index
    }
    export enum Index {
        First,
        Last,
        Next,
        Previous
    }
    export enum SelectionMode {
        Default,
        Selected,
        None
    }
    export enum ValueFormat {
        /**
         * Returns items without format
         */
        Unformatted,
        /**
         * Returns items with format
         */
        Formatted
    }
    // Objects
    export interface Binding {
        document: Document;
        /**
         * Id of the Binding
         */
        id: string;
        type: BindingType;
        /**
         * Adds an event handler to the object using the specified event type.
         * @param eventType - The event type. For binding it can be 'bindingDataChanged' and 'bindingSelectionChanged'
         * @param handler - The name of the handler
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        addHandlerAsync(eventType: EventType, handler: any, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Returns the current selection.
         * @param options - Syntax example: {coercionType: 'matrix,'valueFormat: 'formatted', filterType:'all'}
         *       coercionType: The expected shape of the selection. If not specified returns the bindingType shape. Use Office.CoercionType or text value.
         *       valueFormat: Get data with or without format. Use Office.ValueFormat or text value.
         *       startRow: Used in partial get for table/matrix. Indicates the start row.
         *       startColumn: Used in partial get for table/matrix. Indicates the start column.
         *       rowCount: Used in partial get for table/matrix. Indicates the number of rows from the start row.
         *       columnCount: Used in partial get for table/matrix. Indicates the number of columns from the start column.
         *       filterType: Get the visible or all the data. Useful when filtering data. Use Office.FilterType or text value.
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getDataAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Removes an event handler from the object using the specified event type.
         * @param eventType - The event type. For binding can be 'bindingDataChanged' and 'bindingSelectionChanged'
         * @param options - Syntax example: {handler:eventHandler}
         *       handler: Indicates a specific handler to be removed, if not specified all handlers are removed
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        removeHandlerAsync(eventType: EventType, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Writes the specified data into the current selection.
         * @param data - The data to be set. Either a string or value, 2d array or TableData object
         * @param options - Syntax example: {coercionType:Office.CoercionType.Matrix} or {coercionType: 'matrix'}
         *       coercionType: Explicitly sets the shape of the data object. Use Office.CoercionType or text value. If not supplied is inferred from the data type.
         *       startRow: Used in partial set for table/matrix. Indicates the start row.
         *       startColumn: Used in partial set for table/matrix. Indicates the start column.
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        setDataAsync(data: TableData | any, options?: any, callback?: (result: AsyncResult) => void): void;
    }
    export interface Bindings {
        document: Document;
        /**
         * Creates a binding against a named object in the document
         * @param itemName - Name of the bindable object in the document. For Example 'MyExpenses' table in Excel."
         * @param bindingType - The Office BindingType for the data
         * @param options - Syntax example: {id: "BindingID"}
         *       id: Name of the binding, autogenerated if not supplied.
         *       asyncContext: Object keeping state for the callback
         *       columns: The string[] of the columns involved in the binding
         * @param callback - The optional callback method
         */
        addFromNamedItemAsync(itemName: string, bindingType: BindingType, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Create a binding by prompting the user to make a selection on the document.
         * @param bindingType - The Office BindingType for the data
         * @param options - addFromPromptAsyncOptions- e.g. {promptText: "Please select data", id: "mySales"}
         *       promptText: Greet your users with a friendly word.
         *       asyncContext: Object keeping state for the callback
         *       id: Identifier.
         *       sampleData: A TableData that gives sample table in the Dialog.TableData.Headers is [][] of string.
         * @param callback - The optional callback method
         */
        addFromPromptAsync(bindingType: BindingType, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Create a binding based on what the user's current selection.
         * @param bindingType - The Office BindingType for the data
         * @param options - addFromSelectionAsyncOptions- e.g. {id: "BindingID"}
         *       id: Identifier.
         *       asyncContext: Object keeping state for the callback
         *       columns: The string[] of the columns involved in the binding
         *       sampleData: A TableData that gives sample table in the Dialog.TableData.Headers is [][] of string.
         * @param callback - The optional callback method
         */
        addFromSelectionAsync(bindingType: BindingType, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Gets an array with all the binding objects in the document.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getAllAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Retrieves a binding based on its Name
         * @param id - The binding id
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getByIdAsync(id: string, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Removes the binding from the document
         * @param id - The binding id
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        releaseByIdAsync(id: string, options?: any, callback?: (result: AsyncResult) => void): void;
    }
    export interface Context {
        document: Document;
    }
    export interface CustomXmlNode {
        baseName: string;
        namespaceUri: string;
        nodeType: string;
        /**
         * Gets the nodes associated with the xPath expression.
         * @param xPath - The xPath expression
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getNodesAsync(xPath: string, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Gets the node value.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getNodeValueAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Asynchronously gets the text of an XML node in a custom XML part.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getTextAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Gets the node's XML.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getXmlAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Sets the node value.
         * @param value - The value to be set on the node
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        setNodeValueAsync(value: string, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Asynchronously sets the text of an XML node in a custom XML part.
         * @param text - Required. The text value of the XML node.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        setTextAsync(text: string, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Sets the node XML.
         * @param xml - The XML to be set on the node
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        setXmlAsync(xml: string, options?: any, callback?: (result: AsyncResult) => void): void;
    }
    export interface CustomXmlPart {
        builtIn: boolean;
        id: string;
        namespaceManager: CustomXmlPrefixMappings;
        /**
         * Adds an event handler to the object using the specified event type.
         * @param eventType - The event type. For CustomXmlPartNode it can be 'nodeDeleted', 'nodeInserted' or 'nodeReplaced'
         * @param handler - The name of the handler
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        addHandlerAsync(eventType: EventType, handler?: (result: any) => void, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Deletes the Custom XML Part.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        deleteAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Gets the nodes associated with the xPath expression.
         * @param xPath - The xPath expression
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getNodesAsync(xPath: string, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Gets the XML for the Custom XML Part.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getXmlAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Removes an event handler from the object using the specified event type.
         * @param eventType - The event type. For CustomXmlPartNode it can be 'nodeDeleted', 'nodeInserted' or 'nodeReplaced'
         * @param options - Syntax example: {handler:eventHandler}
         *       handler: Indicates a specific handler to be removed, if not specified all handlers are removed
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        removeHandlerAsync(eventType: EventType, options?: any, callback?: (result: AsyncResult) => void): void;
    }
    export interface CustomXmlParts {
        /**
         * Asynchronously adds a new custom XML part to a file.
         * @param xml - The XML to add to the newly created custom XML part.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - A function that is invoked when the callback returns, whose only parameter is of type AsyncResult.
         */
        addAsync(xml: string, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Asynchronously gets the specified custom XML part by its id.
         * @param id - The id of the custom XML part.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - A function that is invoked when the callback returns, whose only parameter is of type AsyncResult.
         */
        getByIdAsync(id: string, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Asynchronously gets the specified custom XML part(s) by its namespace.
         * @param ns  - The namespace to search.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - A function that is invoked when the callback returns, whose only parameter is of type AsyncResult.
         */
        getByNamespaceAsync(ns: string, options?: any, callback?: (result: AsyncResult) => void): void;
    }
    export interface CustomXmlPrefixMappings {
        /**
         * Adds a namespace.
         * @param prefix - The namespace prefix
         * @param ns - The namespace URI
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        addNamespaceAsync(prefix: string, ns: string, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Gets a namespace  with the specified prefix
         * @param prefix - The namespace prefix
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getNamespaceAsync(prefix: string, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Gets a prefix  for  the specified URI
         * @param ns - The namespace URI
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getPrefixAsync(ns: string, options?: any, callback?: (result: AsyncResult) => void): void;
    }
    export interface Document {
        bindings: Bindings;
        customXmlParts: CustomXmlParts;
        mode: DocumentMode;
        settings: Settings;
        url: string;
        /**
         * Adds an event handler for the specified event type.
         * @param eventType - The event type. For document can be 'DocumentSelectionChanged'
         * @param handler - The name of the handler
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        addHandlerAsync(eventType: EventType, handler: any, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Returns the current view of the presentation.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getActiveViewAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Gets the entire file in slices of up to 4MB.
         * @param fileType - The format in which the file will be returned
         * @param options - Syntax example: {sliceSize:1024}
         *       sliceSize: Specifies the desired slice size (in bytes) up to 4MB. If not specified a default slice size of 4MB will be used.
         * @param callback - The optional callback method
         */
        getFileAsync(fileType: FileType, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Gets file properties of the current document.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getFilePropertiesAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Returns the current selection.
         * @param coercionType - The expected shape of the selection.
         * @param options - Syntax example: {valueFormat: 'formatted', filterType:'all'}
         *       valueFormat: Get data with or without format. Use Office.ValueFormat or text value.
         *       filterType: Get the visible or all the data. Useful when filtering data. Use Office.FilterType or text value.
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getSelectedDataAsync(coercionType: CoercionType, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Goes to the specified object or location in the document.
         * @param id - The identifier of the object or location to go to.
         * @param goToType - The type of the location to go to.
         * @param options - Syntax example: {asyncContext:context}
         *       selectionMode: Use Office.SelectionMode or text value.
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        goToByIdAsync(id: string | number, goToType: GoToType, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Removes an event handler for the specified event type.
         * @param eventType - The event type. For document can be 'DocumentSelectionChanged'
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         *       handler: The name of the handler. If not specified all handlers are removed
         * @param callback - The optional callback method
         */
        removeHandlerAsync(eventType: EventType, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Writes the specified data into the current selection.
         * @param data - The data to be set. Either a string or value, 2d array or TableData object
         * @param options - Syntax example: {coercionType:Office.CoercionType.Matrix} or {coercionType: 'matrix'}
         *       coercionType: Explicitly sets the shape of the data object. Use Office.CoercionType or text value. If not supplied is inferred from the data type.
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        setSelectedDataAsync(data: string | TableData | any[][], options?: any, callback?: (result: AsyncResult) => void): void;
    }
    /**
     * Provides information about the document that raised the SelectionChanged event.
     */
    export interface DocumentSelectionChangedEventArgs {
        /**
         * Gets a Document object that represents the document that raised the SelectionChanged event.
         */
        document: Document;
        /**
         * Get an EventType enumeration value that identifies the kind of event that was raised.
         */
        type: EventType;
    }
    export interface File {
        size: number;
        sliceCount: number;
        /**
         * Closes the File.
         * @param callback - The optional callback method
         */
        closeAsync(callback?: (result: AsyncResult) => void): void;
        /**
         * Gets the specified slice.
         * @param sliceIndex - The index of the slice to be retrieved
         * @param callback - The optional callback method
         */
        getSliceAsync(sliceIndex: number, callback?: (result: AsyncResult) => void): void;
    }
    export interface FileProperties {
        /**
         * File's URL
         */
        url: string
    }
    export interface MatrixBinding extends Binding {
        columnCount: number;
        rowCount: number;
    }
    export interface Settings {
        /**
         * Adds an event handler for the object using the specified event type.
         * @param eventType - The event type. For settings can be 'settingsChanged'
         * @param handler - The name of the handler
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        addHandlerAsync(eventType: EventType, handler: any, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Retrieves the setting with the specified name.
         * @param settingName - The name of the setting
         */
        get(name: string): any;
        /**
         * Gets the latest version of the settings object.
         * @param callback - The optional callback method
         */
        refreshAsync(callback?: (result: AsyncResult) => void): void;
        /**
         * Removes the setting with the specified name.
         * @param settingName - The name of the setting
         */
        remove(name: string): void;
        /**
         * Removes an event handler for the specified event type.
         * @param eventType - The event type. For settings can be 'settingsChanged'
         * @param options - Syntax example: {handler:eventHandler}
         *       handler: Indicates a specific handler to be removed, if not specified all handlers are removed
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        removeHandlerAsync(eventType: EventType, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Saves all settings.
         * @param options - Syntax example: {overwriteIfStale:false}
         *       overwriteIfStale: Indicates whether the setting will be replaced if stale.
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        saveAsync(callback?: (result: AsyncResult) => void): void;
        /**
         * Sets a value for the setting with the specified name.
         * @param settingName - The name of the setting
         * @param value - The value for the setting
         */
        set(name: string, value: any): void;
    }
    export interface Slice {
        data: any;
        index: number;
        size: number;
    }
    export interface TableBinding extends Binding {
        columnCount: number;
        hasHeaders: boolean;
        rowCount: number;
        /**
         * Adds the specified columns to the table
         * @param tableData  - A TableData object with the headers and rows
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        addColumnsAsync(tableData: TableData | any[][], options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Adds the specified rows to the table
         * @param rows  - A 2D array with the rows to add
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        addRowsAsync(rows: TableData | any[][], options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Clears the table
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        deleteAllDataValuesAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Clears formatting on the bound table.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        clearFormatsAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Gets the formatting on specified items in the table.
         * @param cellReference - An object literal containing name-value pairs that specify the range of cells to get formatting from.
         * @param formats - An array specifying the format properties to get.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getFormatsAsync(cellReference?: any, formats?: any[], options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Sets formatting on specified items and data in the table.
         * @param formatsInfo - Array elements are themselves three-element arrays:[target, type, formats]
         *       target: The identifier of the item to format. String.
         *       type: The kind of item to format. String.
         *       formats: An object literal containing a list of property name-value pairs that define the formatting to apply.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        setFormatsAsync(formatsInfo?: any[], options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Updates table formatting options on the bound table.
         * @param tableOptions - An object literal containing a list of property name-value pairs that define the table options to apply.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        setTableOptionsAsync(tableOptions: any, options?: any, callback?: (result: AsyncResult) => void): void;
    }
    export class TableData {
        constructor(rows: any[][], headers: any[]);
        constructor();
        headers: any[];
        rows: any[][];
    }
    export enum Table {
        All,
        Data,
        Headers
    }
    export interface TextBinding extends Binding { }
    export enum ProjectProjectFields {
        CurrencyDigits,
        CurrencySymbol,
        CurrencySymbolPosition,
        DurationUnits,
        GUID,
        Finish,
        Start,
        ReadOnly,
        VERSION,
        WorkUnits,
        ProjectServerUrl,
        WSSUrl,
        WSSList
    }
    export enum ProjectResourceFields {
        Accrual,
        ActualCost,
        ActualOvertimeCost,
        ActualOvertimeWork,
        ActualOvertimeWorkProtected,
        ActualWork,
        ActualWorkProtected,
        BaseCalendar,
        Baseline10BudgetCost,
        Baseline10BudgetWork,
        Baseline10Cost,
        Baseline10Work,
        Baseline1BudgetCost,
        Baseline1BudgetWork,
        Baseline1Cost,
        Baseline1Work,
        Baseline2BudgetCost,
        Baseline2BudgetWork,
        Baseline2Cost,
        Baseline2Work,
        Baseline3BudgetCost,
        Baseline3BudgetWork,
        Baseline3Cost,
        Baseline3Work,
        Baseline4BudgetCost,
        Baseline4BudgetWork,
        Baseline4Cost,
        Baseline4Work,
        Baseline5BudgetCost,
        Baseline5BudgetWork,
        Baseline5Cost,
        Baseline5Work,
        Baseline6BudgetCost,
        Baseline6BudgetWork,
        Baseline6Cost,
        Baseline6Work,
        Baseline7BudgetCost,
        Baseline7BudgetWork,
        Baseline7Cost,
        Baseline7Work,
        Baseline8BudgetCost,
        Baseline8BudgetWork,
        Baseline8Cost,
        Baseline8Work,
        Baseline9BudgetCost,
        Baseline9BudgetWork,
        Baseline9Cost,
        Baseline9Work,
        BaselineBudgetCost,
        BaselineBudgetWork,
        BaselineCost,
        BaselineWork,
        BudgetCost,
        BudgetWork,
        ResourceCalendarGUID,
        Code,
        Cost1,
        Cost10,
        Cost2,
        Cost3,
        Cost4,
        Cost5,
        Cost6,
        Cost7,
        Cost8,
        Cost9,
        ResourceCreationDate,
        Date1,
        Date10,
        Date2,
        Date3,
        Date4,
        Date5,
        Date6,
        Date7,
        Date8,
        Date9,
        Duration1,
        Duration10,
        Duration2,
        Duration3,
        Duration4,
        Duration5,
        Duration6,
        Duration7,
        Duration8,
        Duration9,
        Email,
        End,
        Finish1,
        Finish10,
        Finish2,
        Finish3,
        Finish4,
        Finish5,
        Finish6,
        Finish7,
        Finish8,
        Finish9,
        Flag10,
        Flag1,
        Flag11,
        Flag12,
        Flag13,
        Flag14,
        Flag15,
        Flag16,
        Flag17,
        Flag18,
        Flag19,
        Flag2,
        Flag20,
        Flag3,
        Flag4,
        Flag5,
        Flag6,
        Flag7,
        Flag8,
        Flag9,
        Group,
        Units,
        Name,
        Notes,
        Number1,
        Number10,
        Number11,
        Number12,
        Number13,
        Number14,
        Number15,
        Number16,
        Number17,
        Number18,
        Number19,
        Number2,
        Number20,
        Number3,
        Number4,
        Number5,
        Number6,
        Number7,
        Number8,
        Number9,
        OvertimeCost,
        OvertimeRate,
        OvertimeWork,
        PercentWorkComplete,
        CostPerUse,
        Generic,
        OverAllocated,
        RegularWork,
        RemainingCost,
        RemainingOvertimeCost,
        RemainingOvertimeWork,
        RemainingWork,
        ResourceGUID,
        Cost,
        Work,
        Start,
        Start1,
        Start10,
        Start2,
        Start3,
        Start4,
        Start5,
        Start6,
        Start7,
        Start8,
        Start9,
        StandardRate,
        Text1,
        Text10,
        Text11,
        Text12,
        Text13,
        Text14,
        Text15,
        Text16,
        Text17,
        Text18,
        Text19,
        Text2,
        Text20,
        Text21,
        Text22,
        Text23,
        Text24,
        Text25,
        Text26,
        Text27,
        Text28,
        Text29,
        Text3,
        Text30,
        Text4,
        Text5,
        Text6,
        Text7,
        Text8,
        Text9
    }
    export enum ProjectTaskFields {
        ActualCost,
        ActualDuration,
        ActualFinish,
        ActualOvertimeCost,
        ActualOvertimeWork,
        ActualStart,
        ActualWork,
        Text1,
        Text10,
        Finish10,
        Start10,
        Text11,
        Text12,
        Text13,
        Text14,
        Text15,
        Text16,
        Text17,
        Text18,
        Text19,
        Finish1,
        Start1,
        Text2,
        Text20,
        Text21,
        Text22,
        Text23,
        Text24,
        Text25,
        Text26,
        Text27,
        Text28,
        Text29,
        Finish2,
        Start2,
        Text3,
        Text30,
        Finish3,
        Start3,
        Text4,
        Finish4,
        Start4,
        Text5,
        Finish5,
        Start5,
        Text6,
        Finish6,
        Start6,
        Text7,
        Finish7,
        Start7,
        Text8,
        Finish8,
        Start8,
        Text9,
        Finish9,
        Start9,
        Baseline10BudgetCost,
        Baseline10BudgetWork,
        Baseline10Cost,
        Baseline10Duration,
        Baseline10Finish,
        Baseline10FixedCost,
        Baseline10FixedCostAccrual,
        Baseline10Start,
        Baseline10Work,
        Baseline1BudgetCost,
        Baseline1BudgetWork,
        Baseline1Cost,
        Baseline1Duration,
        Baseline1Finish,
        Baseline1FixedCost,
        Baseline1FixedCostAccrual,
        Baseline1Start,
        Baseline1Work,
        Baseline2BudgetCost,
        Baseline2BudgetWork,
        Baseline2Cost,
        Baseline2Duration,
        Baseline2Finish,
        Baseline2FixedCost,
        Baseline2FixedCostAccrual,
        Baseline2Start,
        Baseline2Work,
        Baseline3BudgetCost,
        Baseline3BudgetWork,
        Baseline3Cost,
        Baseline3Duration,
        Baseline3Finish,
        Baseline3FixedCost,
        Baseline3FixedCostAccrual,
        Basline3Start,
        Baseline3Work,
        Baseline4BudgetCost,
        Baseline4BudgetWork,
        Baseline4Cost,
        Baseline4Duration,
        Baseline4Finish,
        Baseline4FixedCost,
        Baseline4FixedCostAccrual,
        Baseline4Start,
        Baseline4Work,
        Baseline5BudgetCost,
        Baseline5BudgetWork,
        Baseline5Cost,
        Baseline5Duration,
        Baseline5Finish,
        Baseline5FixedCost,
        Baseline5FixedCostAccrual,
        Baseline5Start,
        Baseline5Work,
        Baseline6BudgetCost,
        Baseline6BudgetWork,
        Baseline6Cost,
        Baseline6Duration,
        Baseline6Finish,
        Baseline6FixedCost,
        Baseline6FixedCostAccrual,
        Baseline6Start,
        Baseline6Work,
        Baseline7BudgetCost,
        Baseline7BudgetWork,
        Baseline7Cost,
        Baseline7Duration,
        Baseline7Finish,
        Baseline7FixedCost,
        Baseline7FixedCostAccrual,
        Baseline7Start,
        Baseline7Work,
        Baseline8BudgetCost,
        Baseline8BudgetWork,
        Baseline8Cost,
        Baseline8Duration,
        Baseline8Finish,
        Baseline8FixedCost,
        Baseline8FixedCostAccrual,
        Baseline8Start,
        Baseline8Work,
        Baseline9BudgetCost,
        Baseline9BudgetWork,
        Baseline9Cost,
        Baseline9Duration,
        Baseline9Finish,
        Baseline9FixedCost,
        Baseline9FixedCostAccrual,
        Baseline9Start,
        Baseline9Work,
        BaselineBudgetCost,
        BaselineBudgetWork,
        BaselineCost,
        BaselineDuration,
        BaselineFinish,
        BaselineFixedCost,
        BaselineFixedCostAccrual,
        BaselineStart,
        BaselineWork,
        BudgetCost,
        BudgetFixedCost,
        BudgetFixedWork,
        BudgetWork,
        TaskCalendarGUID,
        ConstraintDate,
        ConstraintType,
        Cost1,
        Cost10,
        Cost2,
        Cost3,
        Cost4,
        Cost5,
        Cost6,
        Cost7,
        Cost8,
        Cost9,
        Date1,
        Date10,
        Date2,
        Date3,
        Date4,
        Date5,
        Date6,
        Date7,
        Date8,
        Date9,
        Deadline,
        Duration1,
        Duration10,
        Duration2,
        Duration3,
        Duration4,
        Duration5,
        Duration6,
        Duration7,
        Duration8,
        Duration9,
        Duration,
        EarnedValueMethod,
        FinishSlack,
        FixedCost,
        FixedCostAccrual,
        Flag10,
        Flag1,
        Flag11,
        Flag12,
        Flag13,
        Flag14,
        Flag15,
        Flag16,
        Flag17,
        Flag18,
        Flag19,
        Flag2,
        Flag20,
        Flag3,
        Flag4,
        Flag5,
        Flag6,
        Flag7,
        Flag8,
        Flag9,
        FreeSlack,
        HasRollupSubTasks,
        ID,
        Name,
        Notes,
        Number1,
        Number10,
        Number11,
        Number12,
        Number13,
        Number14,
        Number15,
        Number16,
        Number17,
        Number18,
        Number19,
        Number2,
        Number20,
        Number3,
        Number4,
        Number5,
        Number6,
        Number7,
        Number8,
        Number9,
        ScheduledDuration,
        ScheduledFinish,
        ScheduledStart,
        OutlineLevel,
        OvertimeCost,
        OvertimeWork,
        PercentComplete,
        PercentWorkComplete,
        Predecessors,
        PreleveledFinish,
        PreleveledStart,
        Priority,
        Active,
        Critical,
        Milestone,
        Overallocated,
        IsRollup,
        Summary,
        RegularWork,
        RemainingCost,
        RemainingDuration,
        RemainingOvertimeCost,
        RemainingWork,
        ResourceNames,
        Cost,
        Finish,
        Start,
        Work,
        StartSlack,
        Status,
        Successors,
        StatusManager,
        TotalSlack,
        TaskGUID,
        Type,
        WBS,
        WBSPREDECESSORS,
        WBSSUCCESSORS,
        WSSID
    }
    export enum ProjectViewTypes {
        Gantt,
        NetworkDiagram,
        TaskDiagram,
        TaskForm,
        TaskSheet,
        ResourceForm,
        ResourceSheet,
        ResourceGraph,
        TeamPlanner,
        TaskDetails,
        TaskNameForm,
        ResourceNames,
        Calendar,
        TaskUsage,
        ResourceUsage,
        Timeline
    }
    // Objects
    export interface Document {
        /**
         * Get Project field (Ex. ProjectWebAccessURL).
         * @param fieldId - Project level fields.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getProjectFieldAsync(fieldId: number, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Get resource field for provided resource Id. (Ex.ResourceName)
         * @param resourceId - Either a string or value of the Resource Id.
         * @param fieldId - Resource Fields.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getResourceFieldAsync(resourceId: string, fieldId: number, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Get the current selected Resource's Id.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getSelectedResourceAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Get the current selected Task's Id.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getSelectedTaskAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Get the current selected View Type (Ex. Gantt) and View Name.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getSelectedViewAsync(options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Get the Task Name, WSS Task Id, and ResourceNames for given taskId.
         * @param taskId - Either a string or value of the Task Id.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getTaskAsync(taskId: string, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Get task field for provided task Id. (Ex. StartDate).
         * @param taskId - Either a string or value of the Task Id.
         * @param fieldId - Task Fields.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getTaskFieldAsync(taskId: string, fieldId: number, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Get the WSS Url and list name for the Tasks List, the MPP is synced too.
         * @param options - Syntax example: {asyncContext:context}
         *       asyncContext: Object keeping state for the callback
         * @param callback - The optional callback method
         */
        getWSSUrlAsync(options?: any, callback?: (result: AsyncResult) => void): void;
    }
}




////////////////////////////////////////////////////////////////
////////////////////// Begin Exchange APIs /////////////////////
////////////////////////////////////////////////////////////////

export declare namespace Office {
    export module MailboxEnums {
        export enum AttachmentType {
            /**
             * The attachment is a file
             */
            File,
            /**
             * The attachment is an Exchange item
             */
            Item
        }
        export enum EntityType {
            /**
             * Specifies that the entity is a meeting suggestion
             */
            MeetingSuggestion,
            /**
             * Specifies that the entity is a task suggestion
             */
            TaskSuggestion,
            /**
             * Specifies that the entity is a postal address
             */
            Address,
            /**
             * Specifies that the entity is SMTP email address
             */
            EmailAddress,
            /**
             * Specifies that the entity is an Internet URL
             */
            Url,
            /**
             * Specifies that the entity is US phone number
             */
            PhoneNumber,
            /**
             * Specifies that the entity is a contact
             */
            Contact
        }
        export enum ItemNotificationMessageType {
            /**
             * The notificationMessage is a progress indicator.
             */
            ProgressIndicator,
            /**
             * The notificationMessage is an informational message.
             */
            InformationalMessage,
            /**
             * The notificationMessage is an error message.
             */
            ErrorMessage
        }
        export enum ItemType {
            /**
             * An email, meeting request, meeting response, or meeting cancellation
             */
            Message,
            /**
             * An appointment item
             */
            Appointment
        }
        export enum ResponseType {
            /**
             * There has been no response from the attendee
             */
            None,
            /**
             * The attendee is the meeting organizer
             */
            Organizer,
            /**
             * The meeting request was tentatively accepted by the attendee
             */
            Tentative,
            /**
             * The meeting request was accepted by the attendee
             */
            Accepted,
            /**
             * The meeting request was declined by the attendee
             */
            Declined
        }
        export enum RecipientType {
            /**
             * Specifies that the recipient is a distribution list containing a list of email addresses
             */
            DistributionList,
            /**
             * Specifies that the recipient is an SMTP email address that is on the Exchange server
             */
            User,
            /**
             * Specifies that the recipient is an SMTP email address that is not on the Exchange server
             */
            ExternalUser,
            /**
             * Specifies that the recipient is not one of the other recipient types
             */
            Other
        }
        export enum RestVersion {
            v1_0,
            v2_0,
            Beta
        }
    }
    export module cast {
        export module item {
            export function toAppointmentCompose(item: Office.Item): Office.AppointmentCompose;
            export function toAppointmentRead(item: Office.Item): Office.AppointmentRead;
            export function toAppointment(item: Office.Item): Office.Appointment;
            export function toMessageCompose(item: Office.Item): Office.MessageCompose;
            export function toMessageRead(item: Office.Item): Office.MessageRead;
            export function toMessage(item: Office.Item): Office.Message;
            export function toItemCompose(item: Office.Item): Office.ItemCompose;
            export function toItemRead(item: Office.Item): Office.ItemRead;
        }
    }
    export interface AsyncContextOptions {
        asyncContext?: any;
    }
    export interface CoercionTypeOptions {
        coercionType?: CoercionType;
    }
    export enum SourceProperty {
        /**
         * The source of the data is from the body of the message.
         */
        Body,
        /**
         * The source of the data is from the subject of the message.
         */
        Subject
    }
    export interface Appointment extends Item {
    }
    export interface AppointmentCompose extends Appointment, ItemCompose {
        end: Time;
        location: Location;
        optionalAttendees: Recipients;
        requiredAttendees: Recipients;
        start: Time;
    }
    export interface AppointmentRead extends Appointment, ItemRead {
        end: Date;
        location: string;
        optionalAttendees: Array<EmailAddressDetails>;
        organizer: EmailAddressDetails;
        requiredAttendees: Array<EmailAddressDetails>;
        resources: EmailAddressDetails;
        start: Date;
    }
    export interface AppointmentForm {
        requiredAttendees: Array<string> | Array<EmailAddressDetails>;
        optionalAttendees: Array<string> | Array<EmailAddressDetails>;
        start: Date;
        end: Date;
        location: string;
        resources: Array<string>;
        subject: string;
        body: string;
    }
    export interface AttachmentDetails {
        attachmentType: Office.MailboxEnums.AttachmentType;
        contentType: string;
        id: string;
        isInline: boolean;
        name: string;
        size: number;
    }
    export interface Body {
        /**
         * Returns the current body in a specified format
         * @param coercionType - The format of the returned body
         * @param callback - optional method to call when the getAsync method returns
         */
        getAsync(coercionType: CoercionType, callback: (result: AsyncResult) => void): void;
        /**
         * Returns the current body in a specified format
         * @param coercionType - The format of the returned body
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - optional method to call when the getAsync method returns
         */
        getAsync(coercionType: CoercionType, options: AsyncContextOptions, callback: (result: AsyncResult) => void): void;

        /*
         * Gets a value that indicates whether the content is in HTML or text format
         * @param tableData  - A TableData object with the headers and rows
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - The optional method to call when the getTypeAsync method returns
         */
        getTypeAsync(options?: AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Adds the specified content to the beginning of the item body
         * @param data - The string to be inserted at the beginning of the body. The string is limited to 1,000,000 characters
         */
        prependAsync(data: string): void;
        /**
         * Adds the specified content to the beginning of the item body
         * @param data - The string to be inserted at the beginning of the body. The string is limited to 1,000,000 characters
         * @param options - Any optional parameters or state data passed to the method
         */
        prependAsync(data: string, options: AsyncContextOptions & CoercionTypeOptions): void;
        /**
         * Adds the specified content to the beginning of the item body
         * @param data - The string to be inserted at the beginning of the body. The string is limited to 1,000,000 characters
         * @param callback - The optional method to call when the string is inserted
         */
        prependAsync(data: string, callback: (result: AsyncResult) => void): void;
        /**
         * Adds the specified content to the beginning of the item body
         * @param data - The string to be inserted at the beginning of the body. The string is limited to 1,000,000 characters
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - The optional method to call when the string is inserted
         */
        prependAsync(data: string, options: AsyncContextOptions & CoercionTypeOptions, callback: (result: AsyncResult) => void): void;

        /**
         * Replaces the entire body with the specified text.
         * @param data - The string that will replace the existing body. The string is limited to 1,000,000 characters
         */
        setAsync(data: string): void;
        /**
         * Replaces the entire body with the specified text.
         * @param data - The string that will replace the existing body. The string is limited to 1,000,000 characters
         * @param options - Any optional parameters or state data passed to the method
         */
        setAsync(data: string, options: AsyncContextOptions & CoercionTypeOptions): void;
        /**
         * Replaces the entire body with the specified text.
         * @param data - The string that will replace the existing body. The string is limited to 1,000,000 characters
         * @param callback - the optional method to call when the body is replaced
         */
        setAsync(data: string, callback: (result: AsyncResult) => void): void;
        /**
         * Replaces the entire body with the specified text.
         * @param data - The string that will replace the existing body. The string is limited to 1,000,000 characters
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - the optional method to call when the body is replaced
         */
        setAsync(data: string, options: AsyncContextOptions & CoercionTypeOptions, callback: (result: AsyncResult) => void): void;

        /**
         * Replaces the selection in the body with the specified text
         * @param data - The string to be inserted at the beginning of the body. The string is limited to 1,000,000 characters
         */
        setSelectedDataAsync(data: string): void;
        /**
         * Replaces the selection in the body with the specified text
         * @param data - The string to be inserted at the beginning of the body. The string is limited to 1,000,000 characters
         * @param options - Any optional parameters or state data passed to the method
         */
        setSelectedDataAsync(data: string, options: AsyncContextOptions & CoercionTypeOptions): void;
        /**
         * Replaces the selection in the body with the specified text
         * @param data - The string to be inserted at the beginning of the body. The string is limited to 1,000,000 characters
         * @param callback - The optional method to call when the string is inserted
         */
        setSelectedDataAsync(data: string, callback: (result: AsyncResult) => void): void;
        /**
         * Replaces the selection in the body with the specified text
         * @param data - The string to be inserted at the beginning of the body. The string is limited to 1,000,000 characters
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - The optional method to call when the string is inserted
         */
        setSelectedDataAsync(data: string, options: AsyncContextOptions & CoercionTypeOptions, callback: (result: AsyncResult) => void): void;

    }
    export interface Contact {
        addresses: Array<string>;
        businessName: string;
        emailAddresses: Array<string>;
        personName: string;
        phoneNumbers: Array<PhoneNumber>;
        urls: Array<string>;
    }
    export interface Context {
        mailbox: Mailbox;
        roamingSettings: RoamingSettings;
    }
    export interface CustomProperties {
        /**
         * Returns the value of the specified custom property
         * @param name - The name of the property to be returned
         */
        get(name: string): any;
        /**
         * Sets the specified property to the specified value
         * @param name - The name of the property to be set
         * @param value - The value of the property to be set
         */
        set(name: string, value: string): void;
        /**
         * Removes the specified property from the custom property collection.
         * @param name - The name of the property to be removed
         */
        remove(name: string): void;
        /**
         * Saves the custom property collection to the server
         * @param callback - The optional callback method
         * @param userContext - Optional variable for any state data that is passed to the saveAsync method
         */
        saveAsync(callback?: (result: AsyncResult) => void, userContext?: any): void;
    }
    export interface Diagnostics {
        hostName: string;
        hostVersion: string;
        OWAView: string;
    }
    export interface EmailAddressDetails {
        emailAddress: string;
        displayName: string;
        appointmentResponse: Office.MailboxEnums.ResponseType;
        recipientType: Office.MailboxEnums.RecipientType;
    }
    export interface EmailUser {
        displayName: string;
        emailAddress: string;
    }
    export interface Entities {
        addresses: Array<string>;
        contacts: Array<Contact>;
        emailAddresses: Array<string>;
        meetingSuggestions: Array<MeetingSuggestion>;
        phoneNumbers: Array<PhoneNumber>;
        taskSuggestions: Array<string>;
        urls: Array<string>;
    }
    export interface Item {
        /**
        * You can cast item with `(Item as Office.[CAST_TYPE])` where CAST_TYPE is one of the following: ItemRead, ItemCompose, Message,
        * MessageRead, MessageCompose, Appointment, AppointmentRead, AppointmentCompose
        */
        __BeSureToCastThisObject__: void;
        body: Body;
        itemType: Office.MailboxEnums.ItemType;
        notificationMessages: NotificationMessages;
        dateTimeCreated: Date;
        /**
         * Asynchronously loads custom properties that are specific to the item and a app for Office
         * @param callback - The optional callback method
         * @param userContext - Optional variable for any state data that is passed to the asynchronous method
         */
        loadCustomPropertiesAsync(callback?: (result: AsyncResult) => void, userContext?: any): void;
    }
    export interface ItemCompose extends Item {
        subject: Subject;
        /**
         * Adds a file to a message as an attachment
         * @param uri - The URI that provides the location of the file to attach to the message. The maximum length is 2048 characters
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters
         */
        addFileAttachmentAsync(uri: string, attachmentName: string): void;
        /**
         * Adds a file to a message as an attachment
         * @param uri - The URI that provides the location of the file to attach to the message. The maximum length is 2048 characters
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters
         * @param options - Any optional parameters or state data passed to the method
         */
        addFileAttachmentAsync(uri: string, attachmentName: string, options: AsyncContextOptions): void;
        /**
         * Adds a file to a message as an attachment
         * @param uri - The URI that provides the location of the file to attach to the message. The maximum length is 2048 characters
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters
         * @param callback - The optional callback method
         */
        addFileAttachmentAsync(uri: string, attachmentName: string, callback: (result: AsyncResult) => void): void;
        /**
         * Adds a file to a message as an attachment
         * @param uri - The URI that provides the location of the file to attach to the message. The maximum length is 2048 characters
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - The optional callback method
         */
        addFileAttachmentAsync(uri: string, attachmentName: string, options: AsyncContextOptions, callback: (result: AsyncResult) => void): void;

        /**
         * Adds an Exchange item, such as a message, as an attachment to the message
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string): void;
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters
         * @param options - Any optional parameters or state data passed to the method
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string, options: AsyncContextOptions): void;
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters
         * @param callback - The optional callback method
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string, callback: (result: AsyncResult) => void): void;
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - The optional callback method
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string, options: AsyncContextOptions, callback: (result: AsyncResult) => void): void;

        /**
         * Closes the current item that is being composed
         *
         * The behaviors of the close method depends on the current state of the item being composed. If the item has unsaved changes, the client
         * prompts the user to save, discard, or close the action.
         *
         * In the Outlook desktop client, if the message is an inline reply, the close method has no effect.
         */
        close(): void;
        /**
         * Asynchronously returns selected data from the subject or body of a message.
         *
         * If there is no selection but the cursor is in the body or the subject, the method returns null for the selected data. If a field other
         * than the body or subject is selected, the method returns the InvalidSelection error
         */
        getSelectedDataAsync(coercionType: CoercionType, callback: (result: AsyncResult) => void): void;
        /**
         * Asynchronously returns selected data from the subject or body of a message.
         *
         * If there is no selection but the cursor is in the body or the subject, the method returns null for the selected data. If a field other
         * than the body or subject is selected, the method returns the InvalidSelection error
         */
        getSelectedDataAsync(coercionType: CoercionType, options: AsyncContextOptions, callback: (result: AsyncResult) => void): void;

        /**
         * Removes an attachment from a message
         * @param attachmentIndex - The index of the attachment to remove. The maximum length of the string is 100 characters
         */
        removeAttachmentAsync(attachmentIndex: string): void;
        /**
         * Removes an attachment from a message
         * @param attachmentIndex - The index of the attachment to remove. The maximum length of the string is 100 characters
         * @param options - Any optional parameters or state data passed to the method
         */
        removeAttachmentAsync(attachmentIndex: string, options: AsyncContextOptions): void;
        /**
         * Removes an attachment from a message
         * @param attachmentIndex - The index of the attachment to remove. The maximum length of the string is 100 characters
         * @param callback - The optional callback method
         */
        removeAttachmentAsync(attachmentIndex: string, callback: (result: AsyncResult) => void): void;
        /**
         * Removes an attachment from a message
         * @param attachmentIndex - The index of the attachment to remove. The maximum length of the string is 100 characters
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - The optional callback method
         */
        removeAttachmentAsync(attachmentIndex: string, options: AsyncContextOptions, callback: (result: AsyncResult) => void): void;

        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or
         * Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.
         */
        saveAsync(): void;
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or
         * Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.
         */
        saveAsync(options: AsyncContextOptions): void;
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or
         * Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.
         */
        saveAsync(callback: (result: AsyncResult) => void): void;
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or
         * Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.
         */
        saveAsync(options: AsyncContextOptions, callback: (result: AsyncResult) => void): void;

        /**
         * Asynchronously inserts data into the body or subject of a message.
         */
        setSelectedDataAsync(data: string): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         */
        setSelectedDataAsync(data: string, options: AsyncContextOptions & CoercionTypeOptions): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         */
        setSelectedDataAsync(data: string, callback: (result: AsyncResult) => void): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         */
        setSelectedDataAsync(data: string, options: AsyncContextOptions & CoercionTypeOptions, callback: (result: AsyncResult) => void): void;

    }
    export interface ItemRead extends Item {
        attachments: Array<AttachmentDetails>;
        itemClass: string;
        itemId: string;
        normalizedSubject: string;
        subject: string;
        /**
         * Displays a reply form that includes the sender and all the recipients of the selected message
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *  OR
         * An object that contains body or attachment data and a callback function
         */
        displayReplyAllForm(formData: string | ReplyFormData): void;
        /**
         * Displays a reply form that includes only the sender of the selected message
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *  OR
         * An object that contains body or attachment data and a callback function
         */
        displayReplyForm(formData: string | ReplyFormData): void;
        /**
         * Gets the entities found in the selected item
         */
        getEntities(): Entities;
        /**
         * Gets an array of entities of the specified entity type found in an message
         * @param entityType - One of the EntityType enumeration values
         */
        getEntitiesByType(entityType: Office.MailboxEnums.EntityType): Array<(string | Contact | MeetingSuggestion | PhoneNumber | TaskSuggestion)>;
        /**
         * Returns well-known entities that pass the named filter defined in the manifest XML file
         * @param name - The name of the ItemHasKnownEntity rule element that defines the filter to match
         */
        getFilteredEntitiesByName(name: string): Array<(string | Contact | MeetingSuggestion | PhoneNumber | TaskSuggestion)>;
        /**
         * Returns string values in the currently selected message object that match the regular expressions defined in the manifest XML file
         */
        getRegExMatches(): any;
        /**
         * Returns string values that match the named regular expression defined in the manifest XML file
         */
        getRegExMatchesByName(name: string): Array<string>;
    }
    export interface LocalClientTime {
        month: number;
        date: number;
        year: number;
        hours: number;
        minutes: number;
        seconds: number;
        milliseconds: number;
        timezoneOffset: number;
    }
    export interface Location {
        /**
         * Begins an asynchronous request for the location of an appointment
         * @param callback - The optional method to call when the string is inserted
         */
        getAsync(callback: (result: AsyncResult) => void): void;
        /**
         * Begins an asynchronous request for the location of an appointment
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - The optional method to call when the string is inserted
         */
        getAsync(options: AsyncContextOptions, callback: (result: AsyncResult) => void): void;

        /**
         * Begins an asynchronous request to set the location of an appointment
         * @param data - The location of the appointment. The string is limited to 255 characters
         */
        setAsync(location: string): void;
        /**
         * Begins an asynchronous request to set the location of an appointment
         * @param data - The location of the appointment. The string is limited to 255 characters
         * @param options - Any optional parameters or state data passed to the method
         */
        setAsync(location: string, options: AsyncContextOptions): void;
        /**
         * Begins an asynchronous request to set the location of an appointment
         * @param data - The location of the appointment. The string is limited to 255 characters
         * @param callback - The optional method to call when the location is set
         */
        setAsync(location: string, callback: (result: AsyncResult) => void): void;
        /**
         * Begins an asynchronous request to set the location of an appointment
         * @param data - The location of the appointment. The string is limited to 255 characters
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - The optional method to call when the location is set
         */
        setAsync(location: string, options: AsyncContextOptions, callback: (result: AsyncResult) => void): void;

    }
    export interface Mailbox {
        diagnostics: Diagnostics;
        ewsUrl: string;
        item: Item;
        userProfile: UserProfile;
        /**
         * Adds an event handler for a supported event
         * @param eventType - The event that should invoke the handler
         * @param handler - The function to handle the event
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - The optional method to call when the handler is added
         */
        addHandlerAsync(eventType: Office.EventType, handler: (type: Office.EventType) => void, options?: any, callback?: (result: AsyncResult) => void): void;
        /**
         * Converts an item ID formatted for REST into EWS format.
         * @param itemId - An item ID formatted for the Outlook REST APIs
         * @param restVersion - A value indicating the version of the Outlook REST API used to retrieve the item ID
         */
        convertToEwsId(itemId: string, restVersion: Office.MailboxEnums.RestVersion): string;
        /**
         * Gets a Date object from a dictionary containing time information
         * @param timeValue - A Date object
         */
        convertToLocalClientTime(timeValue: Date): LocalClientTime;
        /**
         * Converts an item ID formatted for EWS into REST format.
         * @param itemId - An item ID formatted for the Outlook EWS APIs
         * @param restVersion - A value indicating the version of the Outlook REST API that the converted ID will be used with
         */
        convertToRestId(itemId: string, restVersion: Office.MailboxEnums.RestVersion): string;
        /**
         * Gets a dictionary containing time information in local client time
         * @param input - A dictionary containing a date. The dictionary should contain the following fields: year, month, date, hours, minutes, seconds, time zone, time zone offset
         */
        convertToUtcClientTime(input: LocalClientTime): Date;
        /**
         * Displays an existing calendar appointment
         * @param itemId - The Exchange Web Services (EWS) identifier for an existing calendar appointment
         */
        displayAppointmentForm(itemId: string): void;
        /**
         * Displays an existing message
         * @param itemId - The Exchange Web Services (EWS) identifier for an existing message
         */
        displayMessageForm(itemId: string): void;
        /**
         * Displays a form for creating a new calendar appointment
         * @param parameters - A dictionary of parameters describing the new appointment.
         */
        displayNewAppointmentForm(parameters?: AppointmentForm): void;
        /**
         * Displays a new message form
         * WARNING: This api is not officially released, and may not work on all platforms
         * @param options - A dictionary containing all values to be filled in for the user in the new form
         */
        displayNewMessageForm(options?: any): void;
        /**
         * Gets a string that contains a token used to get an attachment or item from an Exchange Server
         * @param callback - The optional method to call when the string is inserted
         * @param userContext - Optional variable for any state data that is passed to the asynchronous method
         */
        getCallbackTokenAsync(callback?: (result: AsyncResult) => void, userContext?: any): void;
        /**
         * Gets a token identifying the user and the app for Office
         * @param callback - The optional method to call when the string is inserted
         * @param userContext - Optional variable for any state data that is passed to the asynchronous method
         */
        getUserIdentityTokenAsync(callback?: (result: AsyncResult) => void, userContext?: any): void;
        /**
         * Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox
         * @param data - The EWS request
         * @param callback - The optional method to call when the string is inserted
         * @param userContext - Optional variable for any state data that is passed to the asynchronous method
         */
        makeEwsRequestAsync(data: any, callback?: (result: AsyncResult) => void, userContext?: any): void;
    }
    export interface Message extends Item {
        conversationId: string;
    }
    export interface MessageCompose extends Message, ItemCompose {
        bcc: Recipients;
        cc: Recipients;
        to: Recipients;
    }
    export interface MessageRead extends Message, ItemRead {
        cc: Array<EmailAddressDetails>;
        from: EmailAddressDetails;
        internetMessageId: string;
        sender: EmailAddressDetails;
        to: Array<EmailAddressDetails>;
    }
    export interface MeetingSuggestion {
        attendees: Array<EmailUser>;
        end: string;
        location: string;
        meetingstring: string;
        start: string;
        subject: string;
    }
    export interface NotificationMessageDetails {
        key?: string;
        type: Office.MailboxEnums.ItemNotificationMessageType;
        icon?: string;
        message: string;
        persistent?: Boolean;
    }
    export interface NotificationMessages {
        /**
         * Adds a notification to an item
         * @param key - A developer-specified key used to refrence this notification message. Developers can use it to modify this message later.
         * @param JSONmessage - A JSON object that contains the notification message to be added to this item
         */
        addAsync(key: string, JSONmessage: NotificationMessageDetails): void;
        /**
         * Adds a notification to an item
         * @param key - A developer-specified key used to refrence this notification message. Developers can use it to modify this message later.
         * @param JSONmessage - A JSON object that contains the notification message to be added to this item
         * @param options - Any optional parameters or state data passed to the method
         */
        addAsync(key: string, JSONmessage: NotificationMessageDetails, options: AsyncContextOptions): void;
        /**
         * Adds a notification to an item
         * @param key - A developer-specified key used to refrence this notification message. Developers can use it to modify this message later.
         * @param JSONmessage - A JSON object that contains the notification message to be added to this item
         * @param callback - The optional callback method
         */
        addAsync(key: string, JSONmessage: NotificationMessageDetails, callback: (result: AsyncResult) => void): void;
        /**
         * Adds a notification to an item
         * @param key - A developer-specified key used to refrence this notification message. Developers can use it to modify this message later.
         * @param JSONmessage - A JSON object that contains the notification message to be added to this item
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - The optional callback method
         */
        addAsync(key: string, JSONmessage: NotificationMessageDetails, options: AsyncContextOptions, callback: (result: AsyncResult) => void): void;

        /**
         * Returns all keys and messages for an item.
         * @param callback - The optional callback method
         */
        getAllAsync(callback: (result: AsyncResult) => void): void;
        /**
         * Returns all keys and messages for an item.
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - The optional callback method
         */
        getAllAsync(options: AsyncContextOptions, callback: (result: AsyncResult) => void): void;

        /**
         * Removes a notification message for an item.
         * @param key - The key for the notification message to remove
         */
        removeAsync(key: string): void;
        /**
         * Removes a notification message for an item.
         * @param key - The key for the notification message to remove
         * @param options - Any optional parameters or state data passed to the method
         */
        removeAsync(key: string, options: AsyncContextOptions): void;
        /**
         * Removes a notification message for an item.
         * @param key - The key for the notification message to remove
         * @param callback - The optional callback method
         */
        removeAsync(key: string, callback: (result: AsyncResult) => void): void;
        /**
         * Removes a notification message for an item.
         * @param key - The key for the notification message to remove
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - The optional callback method
         */
        removeAsync(key: string, options: AsyncContextOptions, callback: (result: AsyncResult) => void): void;

        /**
         * Replaces a notification message that has a given key with another message
         * @param key - The key for the notification message to replace.
         * @param JSONmessage - A JSON object that contains the new notification message to replace the existing message
         */
        replaceAsync(key: string, JSONmessage: NotificationMessageDetails): void;
        /**
         * Replaces a notification message that has a given key with another message
         * @param key - The key for the notification message to replace.
         * @param JSONmessage - A JSON object that contains the new notification message to replace the existing message
         * @param options - Any optional parameters or state data passed to the method
         */
        replaceAsync(key: string, JSONmessage: NotificationMessageDetails, options: AsyncContextOptions): void;
        /**
         * Replaces a notification message that has a given key with another message
         * @param key - The key for the notification message to replace.
         * @param JSONmessage - A JSON object that contains the new notification message to replace the existing message
         * @param callback - The optional callback method
         */
        replaceAsync(key: string, JSONmessage: NotificationMessageDetails, callback: (result: AsyncResult) => void): void;
        /**
         * Replaces a notification message that has a given key with another message
         * @param key - The key for the notification message to replace.
         * @param JSONmessage - A JSON object that contains the new notification message to replace the existing message
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - The optional callback method
         */
        replaceAsync(key: string, JSONmessage: NotificationMessageDetails, options: AsyncContextOptions, callback: (result: AsyncResult) => void): void;

    }
    export interface PhoneNumber {
        phoneString: string;
        originalPhoneString: string;
        type: string;
    }
    export interface Recipients {
        /**
         * Begins an asynchronous request to add a recipient list to an appointment or message
         * @param recipients - The recipients to add to the recipients list
         */
        addAsync(recipients: Array<string | EmailUser | EmailAddressDetails>): void;
        /**
         * Begins an asynchronous request to add a recipient list to an appointment or message
         * @param recipients - The recipients to add to the recipients list
         * @param options - Any optional parameters or state data passed to the method
         */
        addAsync(recipients: Array<string | EmailUser | EmailAddressDetails>, options: AsyncContextOptions): void;
        /**
         * Begins an asynchronous request to add a recipient list to an appointment or message
         * @param recipients - The recipients to add to the recipients list
         * @param callback - The optional method to call when the string is inserted
         */
        addAsync(recipients: Array<string | EmailUser | EmailAddressDetails>, callback: (result: AsyncResult) => void): void;
        /**
         * Begins an asynchronous request to add a recipient list to an appointment or message
         * @param recipients - The recipients to add to the recipients list
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - The optional method to call when the string is inserted
         */
        addAsync(recipients: Array<string | EmailUser | EmailAddressDetails>, options: AsyncContextOptions, callback: (result: AsyncResult) => void): void;
        /**
         * Begins an asynchronous request to get the recipient list for an appointment or message
         * @param callback - The optional method to call when the string is inserted
         */
        getAsync(callback: (result: AsyncResult) => void): void;
        /**
         * Begins an asynchronous request to get the recipient list for an appointment or message
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - The optional method to call when the string is inserted
         */
        getAsync(options: AsyncContextOptions, callback: (result: AsyncResult) => void): void;

        /**
         * Begins an asynchronous request to set the recipient list for an appointment or message
         * @param recipients - The recipients to add to the recipients list
         */
        setAsync(recipients: Array<string | EmailUser | EmailAddressDetails>): void;
        /**
         * Begins an asynchronous request to set the recipient list for an appointment or message
         * @param recipients - The recipients to add to the recipients list
         * @param options - Any optional parameters or state data passed to the method
         */
        setAsync(recipients: Array<string | EmailUser | EmailAddressDetails>, options: AsyncContextOptions): void;
        /**
         * Begins an asynchronous request to set the recipient list for an appointment or message
         * @param recipients - The recipients to add to the recipients list
         * @param callback - The optional method to call when the string is inserted
         */
        setAsync(recipients: Array<string | EmailUser | EmailAddressDetails>, callback: (result: AsyncResult) => void): void;
        /**
         * Begins an asynchronous request to set the recipient list for an appointment or message
         * @param recipients - The recipients to add to the recipients list
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - The optional method to call when the string is inserted
         */
        setAsync(recipients: Array<string | EmailUser | EmailAddressDetails>, options: AsyncContextOptions, callback: (result: AsyncResult) => void): void;

    }
    export interface ReplyFormAttachment {
        type: string;
        name: string;
        url?: string;
        itemId?: string;
    }
    export interface ReplyFormData {
        htmlBody?: string;
        attachments?: Array<ReplyFormAttachment>;
        callback?: (result: AsyncResult) => void;
    }
    export interface RoamingSettings {
        /**
         * Retrieves the specified setting
         * @param name - The case-sensitive name of the setting to retrieve
         */
        get(name: string): any;
        /**
         * Removes the specified setting
         * @param name - The case-sensitive name of the setting to remove
         */
        remove(name: string): void;
        /**
         * Saves the settings
         * @param callback - A function that is invoked when the callback returns, whose only parameter is of type AsyncResult
         */
        saveAsync(callback?: (result: AsyncResult) => void): void;
        /**
         * Sets or creates the specified setting
         * @param name - The case-sensitive name of the setting to set or create
         * @param value - Specifies the value to be stored
         */
        set(name: string, value: any): void;
    }
    export interface Subject {
        /**
         * Begins an asynchronous request to get the subject of an appointment or message
         * @param callback - The optional method to call when the string is inserted
         */
        getAsync(callback: (result: AsyncResult) => void): void;
        /**
         * Begins an asynchronous request to get the subject of an appointment or message
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - The optional method to call when the string is inserted
         */
        getAsync(options: AsyncContextOptions, callback: (result: AsyncResult) => void): void;

        /**
         * Begins an asynchronous call to set the subject of an appointment or message
         * @param data - The subject of the appointment. The string is limited to 255 characters
         */
        setAsync(data: string): void;
        /**
         * Begins an asynchronous call to set the subject of an appointment or message
         * @param data - The subject of the appointment. The string is limited to 255 characters
         * @param options - Any optional parameters or state data passed to the method
         */
        setAsync(data: string, options: AsyncContextOptions): void;
        /**
         * Begins an asynchronous call to set the subject of an appointment or message
         * @param data - The subject of the appointment. The string is limited to 255 characters
         * @param callback - The optional method to call when the string is inserted
         */
        setAsync(data: string, callback: (result: AsyncResult) => void): void;
        /**
         * Begins an asynchronous call to set the subject of an appointment or message
         * @param data - The subject of the appointment. The string is limited to 255 characters
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - The optional method to call when the string is inserted
         */
        setAsync(data: string, options: AsyncContextOptions, callback: (result: AsyncResult) => void): void;

    }
    export interface TaskSuggestion {
        assignees: Array<EmailUser>;
        taskString: string;
    }
    export interface Time {
        /**
         * Begins an asynchronous request to get the start or end time
         * @param callback - The optional method to call when the string is inserted
         */
        getAsync(callback: (result: AsyncResult) => void): void;
        /**
         * Begins an asynchronous request to get the start or end time
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - The optional method to call when the string is inserted
         */
        getAsync(options: AsyncContextOptions, callback: (result: AsyncResult) => void): void;

        /**
         * Begins an asynchronous request to set the start or end time
         * @param dateTime - A date-time object in Coordinated Universal Time (UTC)
         */
        setAsync(dateTime: Date): void;
        /**
         * Begins an asynchronous request to set the start or end time
         * @param dateTime - A date-time object in Coordinated Universal Time (UTC)
         * @param options - Any optional parameters or state data passed to the method
         */
        setAsync(dateTime: Date, options: AsyncContextOptions): void;
        /**
         * Begins an asynchronous request to set the start or end time
         * @param dateTime - A date-time object in Coordinated Universal Time (UTC)
         * @param callback - The optional method to call when the string is inserted
         */
        setAsync(dateTime: Date, callback: (result: AsyncResult) => void): void;
        /**
         * Begins an asynchronous request to set the start or end time
         * @param dateTime - A date-time object in Coordinated Universal Time (UTC)
         * @param options - Any optional parameters or state data passed to the method
         * @param callback - The optional method to call when the string is inserted
         */
        setAsync(dateTime: Date, options: AsyncContextOptions, callback: (result: AsyncResult) => void): void;

    }
    export interface UserProfile {
        displayName: string;
        emailAddress: string;
        timeZone: string;
    }
}


////////////////////////////////////////////////////////////////
/////////////////////// End Exchange APIs //////////////////////
////////////////////////////////////////////////////////////////



///////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////
//////////////// Begin OfficeExtension runtime /////////////////
////////////////////////////////////////////////////////////////

export declare namespace OfficeExtension {
    /** An abstract proxy object that represents an object in an Office document. You create proxy objects from the context (or from other proxy objects), add commands to a queue to act on the object, and then synchronize the proxy object state with the document by calling "context.sync()". */
    export class ClientObject {
        /** The request context associated with the object */
        context: ClientRequestContext;
        /** Returns a boolean value for whether the corresponding object is a null object. You must call "context.sync()" before reading the isNullObject property. */
        isNullObject: boolean;
    }
}
export declare namespace OfficeExtension {
    export interface LoadOption {
        select?: string | string[];
        expand?: string | string[];
        top?: number;
        skip?: number;
    }
    /** An abstract RequestContext object that facilitates requests to the host Office application. The "Excel.run" and "Word.run" methods provide a request context. */
    export class ClientRequestContext {
        constructor(url?: string);

        /** Collection of objects that are tracked for automatic adjustments based on surrounding changes in the document. */
        trackedObjects: TrackedObjects;

        /** Request headers */
        requestHeaders: { [name: string]: string };

        /** Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties. */
        load(object: ClientObject, option?: string | string[] | LoadOption): void;

		/**
		* Queues up a command to recursively load the specified properties of the object and its navigation properties.
		* You must call "context.sync()" before reading the properties.
		*
		* @param object - The object to be loaded.
		* @param options - The key-value pairing of load options for the types, such as { "Workbook": "worksheets,tables",  "Worksheet": "tables",  "Tables": "name" }
		* @param maxDepth - The maximum recursive depth.
		*/
        loadRecursive(object: ClientObject, options: { [typeName: string]: string | string[] | LoadOption }, maxDepth?: number): void;

        /** Adds a trace message to the queue. If the promise returned by "context.sync()" is rejected due to an error, this adds a ".traceMessages" array to the OfficeExtension.Error object, containing all trace messages that were executed. These messages can help you monitor the program execution sequence and detect the cause of the error. */
        trace(message: string): void;

        /** Synchronizes the state between JavaScript proxy objects and the Office document, by executing instructions queued on the request context and retrieving properties of loaded Office objects for use in your code.�This method returns a promise, which is resolved when the synchronization is complete. */
        sync<T>(passThroughValue?: T): IPromise<T>;
    }

    export interface EmbeddedOptions {
		sessionKey?: string,
		container?: HTMLElement,
		id?: string;
		timeoutInMilliseconds?: number;
		height?: string;
		width?: string;
    }

    export class EmbeddedSession {
        constructor(url: string, options?: EmbeddedOptions);
        public init(): IPromise<any>;
    }
}

export declare namespace OfficeExtension {
    /** Contains the result for methods that return primitive types. The object's value property is retrieved from the document after "context.sync()" is invoked. */
    export class ClientResult<T> {
        /** The value of the result that is retrieved from the document after "context.sync()" is invoked. */
        value: T;
    }
}
export declare namespace OfficeExtension {
    export interface DebugInfo {
        /** Error code string, such as "InvalidArgument". */
        code: string;
        /** The error message passed through from the host Office application. */
        message: string;
        /** Inner error, if applicable. */
        innerError?: DebugInfo | string;

        /** The object type and property or method name (or similar information), if available. */
        errorLocation?: string
    }

    /** The error object returned by "context.sync()", if a promise is rejected due to an error while processing the request. */
    export class Error {
        /** Error name: "OfficeExtension.Error".*/
        name: string;
        /** The error message passed through from the host Office application. */
        message: string;
        /** Stack trace, if applicable. */
        stack: string;
        /** Error code string, such as "InvalidArgument". */
        code: string;
        /** Trace messages (if any) that were added via a "context.trace()" invocation before calling "context.sync()". If there was an error, this contains all trace messages that were executed before the error occurred. These messages can help you monitor the program execution sequence and detect the case of the error. */
        traceMessages: Array<string>;
        /** Debug info (useful for detailed logging of the error, i.e., via JSON.stringify(...)). */
        debugInfo: DebugInfo;
        /** Inner error, if applicable. */
        innerError: Error;
    }
}
export declare namespace OfficeExtension {
    export class ErrorCodes {
        public static accessDenied: string;
        public static generalException: string;
        public static activityLimitReached: string;
        public static invalidObjectPath: string;
        public static propertyNotLoaded: string;
        public static valueNotLoaded: string;
        public static invalidRequestContext: string;
        public static invalidArgument: string;
        public static runMustReturnPromise: string;
        public static cannotRegisterEvent: string;
        public static apiNotFound: string;
        public static connectionFailure: string;
    }
}
export declare namespace OfficeExtension {
    /** An IPromise object that represents a deferred interaction with the host Office application. */
    export interface IPromise<R> {
		/**
		 * This method will be called once the previous promise has been resolved.
		 * Both the onFulfilled on onRejected callbacks are optional.
		 * If either or both are omitted, the next onFulfilled/onRejected in the chain will be called called.

		 * @returns A new promise for the value or error that was returned from onFulfilled/onRejected.
		 */
        then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => IPromise<U>): IPromise<U>;

		/**
		 * This method will be called once the previous promise has been resolved.
		 * Both the onFulfilled on onRejected callbacks are optional.
		 * If either or both are omitted, the next onFulfilled/onRejected in the chain will be called called.

		 * @returns A new promise for the value or error that was returned from onFulfilled/onRejected.
		 */
        then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => U): IPromise<U>;

		/**
		 * This method will be called once the previous promise has been resolved.
		 * Both the onFulfilled on onRejected callbacks are optional.
		 * If either or both are omitted, the next onFulfilled/onRejected in the chain will be called called.

		 * @returns A new promise for the value or error that was returned from onFulfilled/onRejected.
		 */
        then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => void): IPromise<U>;

		/**
		 * This method will be called once the previous promise has been resolved.
		 * Both the onFulfilled on onRejected callbacks are optional.
		 * If either or both are omitted, the next onFulfilled/onRejected in the chain will be called called.

		 * @returns A new promise for the value or error that was returned from onFulfilled/onRejected.
		 */
        then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => IPromise<U>): IPromise<U>;

		/**
		 * This method will be called once the previous promise has been resolved.
		 * Both the onFulfilled on onRejected callbacks are optional.
		 * If either or both are omitted, the next onFulfilled/onRejected in the chain will be called called.

		 * @returns A new promise for the value or error that was returned from onFulfilled/onRejected.
		 */
        then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => U): IPromise<U>;

		/**
		 * This method will be called once the previous promise has been resolved.
		 * Both the onFulfilled on onRejected callbacks are optional.
		 * If either or both are omitted, the next onFulfilled/onRejected in the chain will be called called.

		 * @returns A new promise for the value or error that was returned from onFulfilled/onRejected.
		 */
        then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => void): IPromise<U>;


		/**
		 * Catches failures or exceptions from actions within the promise, or from an unhandled exception earlier in the call stack.
		 * @param onRejected - function to be called if or when the promise rejects.
		 */
        catch<U>(onRejected?: (error: any) => IPromise<U>): IPromise<U>;

		/**
		 * Catches failures or exceptions from actions within the promise, or from an unhandled exception earlier in the call stack.
		 * @param onRejected - function to be called if or when the promise rejects.
		 */
        catch<U>(onRejected?: (error: any) => U): IPromise<U>;

		/**
		 * Catches failures or exceptions from actions within the promise, or from an unhandled exception earlier in the call stack.
		 * @param onRejected - function to be called if or when the promise rejects.
		 */
        catch<U>(onRejected?: (error: any) => void): IPromise<U>;
    }

    /** An Promise object that represents a deferred interaction with the host Office application. The publically-consumable OfficeExtension.Promise is available starting in ExcelApi 1.2 and WordApi 1.2. Promises can be chained via ".then", and errors can be caught via ".catch". Remember to always use a ".catch" on the outer promise, and to return intermediary promises so as not to break the promise chain. When a "native" Promise implementation is available, OfficeExtension.Promise will switch to use the native Promise instead. */
    export class Promise<R> implements IPromise<R>
    {
		/**
		 * Creates a new promise based on a function that accepts resolve and reject handlers.
		 */
        constructor(func: (resolve: (value?: R | IPromise<R>) => void, reject: (error?: any) => void) => void);

		/**
		 * Creates a promise that resolves when all of the child promises resolve.
		 */
        static all<U>(promises: OfficeExtension.IPromise<U>[]): IPromise<U[]>;

		/**
		 * Creates a promise that is resolved.
		 */
        static resolve<U>(value: U): IPromise<U>;

		/**
		 * Creates a promise that is rejected.
		 */
        static reject<U>(error: any): IPromise<U>;

		/* This method will be called once the previous promise has been resolved.
		 * Both the onFulfilled on onRejected callbacks are optional.
		 * If either or both are omitted, the next onFulfilled/onRejected in the chain will be called called.

		 * @returns A new promise for the value or error that was returned from onFulfilled/onRejected.
		 */
        then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => IPromise<U>): IPromise<U>;

		/**
		 * This method will be called once the previous promise has been resolved.
		 * Both the onFulfilled on onRejected callbacks are optional.
		 * If either or both are omitted, the next onFulfilled/onRejected in the chain will be called called.

		 * @returns A new promise for the value or error that was returned from onFulfilled/onRejected.
		 */
        then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => U): IPromise<U>;

		/**
		 * This method will be called once the previous promise has been resolved.
		 * Both the onFulfilled on onRejected callbacks are optional.
		 * If either or both are omitted, the next onFulfilled/onRejected in the chain will be called called.

		 * @returns A new promise for the value or error that was returned from onFulfilled/onRejected.
		 */
        then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => void): IPromise<U>;

		/**
		 * This method will be called once the previous promise has been resolved.
		 * Both the onFulfilled on onRejected callbacks are optional.
		 * If either or both are omitted, the next onFulfilled/onRejected in the chain will be called called.

		 * @returns A new promise for the value or error that was returned from onFulfilled/onRejected.
		 */
        then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => IPromise<U>): IPromise<U>;

		/**
		 * This method will be called once the previous promise has been resolved.
		 * Both the onFulfilled on onRejected callbacks are optional.
		 * If either or both are omitted, the next onFulfilled/onRejected in the chain will be called called.

		 * @returns A new promise for the value or error that was returned from onFulfilled/onRejected.
		 */
        then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => U): IPromise<U>;

		/**
		 * This method will be called once the previous promise has been resolved.
		 * Both the onFulfilled on onRejected callbacks are optional.
		 * If either or both are omitted, the next onFulfilled/onRejected in the chain will be called called.

		 * @returns A new promise for the value or error that was returned from onFulfilled/onRejected.
		 */
        then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => void): IPromise<U>;


		/**
		 * Catches failures or exceptions from actions within the promise, or from an unhandled exception earlier in the call stack.
		 * @param onRejected - function to be called if or when the promise rejects.
		 */
        catch<U>(onRejected?: (error: any) => IPromise<U>): IPromise<U>;

		/**
		 * Catches failures or exceptions from actions within the promise, or from an unhandled exception earlier in the call stack.
		 * @param onRejected - function to be called if or when the promise rejects.
		 */
        catch<U>(onRejected?: (error: any) => U): IPromise<U>;

		/**
		 * Catches failures or exceptions from actions within the promise, or from an unhandled exception earlier in the call stack.
		 * @param onRejected - function to be called if or when the promise rejects.
		 */
        catch<U>(onRejected?: (error: any) => void): IPromise<U>;
    }
}

export declare namespace OfficeExtension {
    /** Collection of tracked objects, contained within a request context. See "context.trackedObjects" for more information. */
    export class TrackedObjects {
        /** Track a new object for automatic adjustment based on surrounding changes in the document. Only some object types require this. If you are using an object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created. */
        add(object: ClientObject): void;
        /** Track a new object for automatic adjustment based on surrounding changes in the document. Only some object types require this. If you are using an object across ".sync" calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you needed to have added the object to the tracked object collection when the object was first created. */
        add(objects: ClientObject[]): void;
        /** Release the memory associated with an object that was previously added to this collection. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect. */
        remove(object: ClientObject): void;
        /** Release the memory associated with an object that was previously added to this collection. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You will need to call "context.sync()" before the memory release takes effect. */
        remove(objects: ClientObject[]): void;
    }
}

export declare namespace OfficeExtension {
    export class EventHandlers<T> {
        constructor(context: ClientRequestContext, parentObject: ClientObject, name: string, eventInfo: EventInfo<T>);
        add(handler: (args: T) => IPromise<any>): EventHandlerResult<T>;
        remove(handler: (args: T) => IPromise<any>): void;
    }

    export class EventHandlerResult<T> {
        constructor(context: ClientRequestContext, handlers: EventHandlers<T>, handler: (args: T) => IPromise<any>);
        remove(): void;
    }

    export interface EventInfo<T> {
        registerFunc: (callback: (args: any) => void) => IPromise<any>;
        unregisterFunc: (callback: (args: any) => void) => IPromise<any>;
        eventArgsTransformFunc: (args: any) => IPromise<T>;
    }
}
export declare namespace OfficeExtension {
	/**
	* Request URL and headers
	*/
    export interface RequestUrlAndHeaderInfo {
        /** Request URL */
        url: string;
        /** Request headers */
        headers?: {
            [name: string]: string;
        };
    }
}



export declare namespace OfficeCore {
    /**
     * [Api set: Experiment 1.1 (PREVIEW)]
     */
    export class FlightingService extends OfficeExtension.ClientObject {
        getFeature(featureName: string, type: string, defaultValue: number | boolean | string, possibleValues?: Array<number> | Array<string> | Array<boolean> | Array<ScopedValue>): OfficeCore.ABType;
        getFeatureGate(featureName: string, scope?: string): OfficeCore.ABType;
        resetOverride(featureName: string): void;
        setOverride(featureName: string, type: string, value: number | boolean | string): void;
        /**
         * Create a new instance of OfficeCore.FlightingService object
         */
        static newObject(context: OfficeExtension.ClientRequestContext): OfficeCore.FlightingService;
        toJSON(): {};
    }
    /**
     *
     * Provides information about the scoped value.
     *
     * [Api set: Experiment 1.1 (PREVIEW)]
     */
    export interface ScopedValue {
        /**
         *
         * Gets the scope.
         *
         * [Api set: Experiment 1.1 (PREVIEW)]
         */
        scope: string;
        /**
         *
         * Gets the value.
         *
         * [Api set: Experiment 1.1 (PREVIEW)]
         */
        value: string | number | boolean;
    }
    /**
     * [Api set: Experiment 1.1 (PREVIEW)]
     */
    export class ABType extends OfficeExtension.ClientObject {
        readonly value: string | number | boolean;
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        load(option?: string | string[] | OfficeExtension.LoadOption): OfficeCore.ABType;
        toJSON(): {
            "value": string | number | boolean;
        };
    }
    /**
     * [Api set: Experiment 1.1 (PREVIEW)]
     */
    export namespace FeatureType {
        var boolean: string;
        var integer: string;
        var string: string;
    }
    export namespace ExperimentErrorCodes {
        var generalException: string;
    }
    export module Interfaces {
    }
}
export declare namespace OfficeCore {
    export class RequestContext extends OfficeExtension.ClientRequestContext {
        constructor(url?: string | OfficeExtension.RequestUrlAndHeaderInfo | any);
        readonly flightingService: FlightingService;
    }
}


////////////////////////////////////////////////////////////////
///////////////// End OfficeExtension runtime //////////////////
////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////



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