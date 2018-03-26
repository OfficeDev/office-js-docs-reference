// Type definitions for Office.js
// Project: http://dev.office.com
// Definitions by: OfficeDev <https://github.com/OfficeDev>, Lance Austin <https://github.com/LanceEA>, Michael Zlatkovsky <https://github.com/Zlatkovsky>, Kim Brandl <https://github.com/kbrandl>, Ricky Kirkham <https://github.com/Rick-Kirkham>
// Definitions: https://github.com/DefinitelyTyped/DefinitelyTyped

/*
office-js
Copyright (c) Microsoft Corporation
*/

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
         * @param tableData - A TableData object with the headers and rows
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
