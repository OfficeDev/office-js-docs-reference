////////////////////////////////////////////////////////////////
////////////////////// Begin Exchange APIs /////////////////////
////////////////////////////////////////////////////////////////

export declare namespace Office {
    export namespace MailboxEnums {
        /**
         * Specifies an attachment's type.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>
         * {@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}
         * </td><td>Compose or read</td></tr></table>
         */
        enum AttachmentType {
            /**
             * The attachment is a file
             */
            File = "file",
            /**
             * The attachment is an Exchange item
             */
            Item = "item",
            /**
             * The attachment is stored in a cloud location, such as OneDrive. The id property of the attachment contains a URL to the file.
             */
            Cloud = "cloud"
        }
        /**
         * Specifies an entity's type.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>
         * {@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}
         * </td><td>Compose or read</td></tr></table>
         */
        enum EntityType {
            /**
             * Specifies that the entity is a meeting suggestion.
             */
            MeetingSuggestion = "meetingSuggestion",
            /**
             * Specifies that the entity is a task suggestion.
             */
            TaskSuggestion = "taskSuggestion",
            /**
             * Specifies that the entity is a postal address.
             */
            Address = "address",
            /**
             * Specifies that the entity is an SMTP email address.
             */
            EmailAddress = "emailAddress",
            /**
             * Specifies that the entity is an Internet URL.
             */
            Url = "url",
            /**
             * Specifies that the entity is a US phone number.
             */
            PhoneNumber = "phoneNumber",
            /**
             * Specifies that the entity is a contact.
             */
            Contact = "contact"
        }
        /**
         * Specifies the notification message type for an appointment or message.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>
         * {@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}
         * </td><td>Compose or read</td></tr></table>
         */
        enum ItemNotificationMessageType {
            /**
             * The notificationMessage is a progress indicator.
             */
            ProgressIndicator = "progressIndicator",
            /**
             * The notificationMessage is an informational message.
             */
            InformationalMessage = "informationalMessage",
            /**
             * The notificationMessage is an error message.
             */
            ErrorMessage = "errorMessage"
        }
        /**
         * Specifies an item's type.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>
         * {@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}
         * </td><td>Compose or read</td></tr></table>
         */
        enum ItemType {
            /**
             * An email, meeting request, meeting response, or meeting cancellation.
             */
            Message = "message",
            /**
             * An appointment item.
             */
            Appointment = "appointment"
        }
        /**
         * Specifies the type of recipient for an appointment.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>
         * {@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}
         * </td><td>Compose or read</td></tr></table>
         */
        enum RecipientType {
            /**
             * Specifies that the recipient is a distribution list containing a list of email addresses.
             */
            DistributionList = "distributionList",
            /**
             * Specifies that the recipient is an SMTP email address that is on the Exchange server.
             */
            User = "user",
            /**
             * Specifies that the recipient is an SMTP email address that is not on the Exchange server.
             */
            ExternalUser = "externalUser",
            /**
             * Specifies that the recipient is not one of the other recipient types.
             */
            Other = "other"
        }
        /**  
         * Specifies the type of response to a meeting invitation.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>
         * {@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}
         * </td><td>Compose or read</td></tr></table>
         */
        enum ResponseType {
            /**
             * There has been no response from the attendee.
             */
            None = "none",
            /**
             * The attendee is the meeting organizer.
             */
            Organizer = "organizer",
            /**
             * The meeting request was tentatively accepted by the attendee.
             */
            Tentative = "tentative",
            /**
             * The meeting request was accepted by the attendee.
             */
            Accepted = "accepted",
            /**
             * The meeting request was declined by the attendee.
             */
            Declined = "declined"
        }
        /**
         * Specifies the version of the REST API that corresponds to a REST-formatted item ID.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>
         * {@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}
         * </td><td>Compose or read</td></tr></table>
         */
        enum RestVersion {
            /**
             * Version 1.0.
             */
            v1_0 = "v1.0",
            /**
             * Version 2.0.
             */
            v2_0 = "v2.0",
            /**
             * Beta.
             */
            Beta = "beta"
        }
    }
    export interface CoercionTypeOptions {
        coercionType?: CoercionType;
    }
    enum SourceProperty {
        /**
         * The source of the data is from the body of the message.
         */
        Body,
        /**
         * The source of the data is from the subject of the message.
         */
        Subject
    }
    /**
     * The AppointmentForm namespace is used to access the currently selected appointment.
     *
     * [Api set: Mailbox 1.0]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>Restricted</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
     */
    export interface AppointmentForm {
        /**
         * Gets an object that provides methods for manipulating the body of an item.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        body: string;
        /**
         * Gets or sets the date and time that the appointment is to end.
         *
         * The end property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the convertToLocalClientTime method to convert the end property value to the client's local date and time.
         *
         * *Read mode*
         *
         * The end property returns a Date object.
         *
         * *Compose mode*
         *
         * The end property returns a Time object.
         *
         * When you use the Time.setAsync method to set the end time, you should use the convertToUtcClientTime method to convert the local time on the client to UTC for the server.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        end: Date;
        /**
        * Gets or sets the location of an appointment.
        *
        * *Read mode*
        *
        * The location property returns a string that contains the location of the appointment.
        *
        * *Compose mode*
        *
        * The location property returns a Location object that provides methods that are used to get and set the location of the appointment.
        *
        * [Api set: Mailbox 1.0]
        *
        * @remarks
        *
        * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
        * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
        */
        location: string;
        /**
        * Provides access to the optional attendees of an event. The type of object and level of access depends on the mode of the current item.
        *
        * *Read mode*
        *
        * The optionalAttendees property returns an array that contains an EmailAddressDetails object for each optional attendee to the meeting.
        *
        * *Compose mode*
        *
        * The optionalAttendees property returns a Recipients object that provides methods to get or update the optional attendees for a meeting.
        *
        * [Api set: Mailbox 1.0]
        *
        * @remarks
        *
        * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
        *
        * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
        */
       optionalAttendees: string[] | EmailAddressDetails[];
       resources: string[];
        /**
         * Provides access to the required attendees of an event. The type of object and level of access depends on the mode of the current item.
         *
         * *Read mode*
         *
         * The requiredAttendees property returns an array that contains an EmailAddressDetails object for each required attendee to the meeting.
         *
         * *Compose mode*
         *
         * The requiredAttendees property returns a Recipients object that provides methods to get or update the required attendees for a meeting.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        requiredAttendees: string[] | EmailAddressDetails[];
        /**
         * Gets or sets the date and time that the appointment is to begin.
         *
         * The start property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the convertToLocalClientTime method to convert the value to the client's local date and time.
         *
         * *Read mode*
         *
         * The start property returns a Date object.
         *
         * *Compose mode*
         *
         * The start property returns a Time object.
         *
         * When you use the Time.setAsync method to set the start time, you should use the convertToUtcClientTime method to convert the local time on the client to UTC for the server.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        start: Date;
        /**
         * Gets or sets the description that appears in the subject field of an item.
         *
         * The subject property gets or sets the entire subject of the item, as sent by the email server.
         *
         * *Read mode*
         *
         * The subject property returns a string. Use the normalizedSubject property to get the subject minus any leading prefixes such as RE: and FW:.
         *
         * *Compose mode*
         *
         * The subject property returns a Subject object that provides methods to get and set the subject.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        subject: string;
    }
    /**
     * Represents an attachment on an item from the server. Read mode only.
     *
     * An array of AttachmentDetail objects is returned as the attachments property of an Appointment or Message object.
     *
     * [Api set: Mailbox 1.0]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
     */
    export interface AttachmentDetails {
        /**
         * Gets a value that indicates the type of an attachment.
         */
        attachmentType: Office.MailboxEnums.AttachmentType;
        /**
         * Gets the MIME content type of the attachment.
         */
        contentType: string;
        /**
         * Gets the Exchange attachment ID of the attachment.
         */
        id: string;
        /**
         * Gets a value that indicates whether the attachment should be displayed in the body of the item.
         */
        isInline: boolean;
        /**
         * Gets the name of the attachment.
         */
        name: string;
        /**
         * Gets the size of the attachment in bytes.
         */
        size: number;
    }
    /**
     * The body object provides methods for adding and updating the content of the message or appointment. It is returned in the body property of the selected item.
     *
     * [Api set: Mailbox 1.1]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
     */
    export interface Body {
        /**
         * Returns the current body in a specified format.
         *
         * This method returns the entire current body in the format specified by coercionType.
         *
         * When working with HTML-formatted bodies, it is important to note that the Body.getAsync and Body.setAsync methods are not idempotent. The value returned from the getAsync method will not necessarily be exactly the same as the value that was passed in the setAsync method previously. The client may modify the value passed to setAsync in order to make it render efficiently with its rendering engine.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         * 
         * In addition to the main signature, this method also has this signature:
         * 
         * `getAsync(coerciontype: Office.CoercionType, callback: (result: AsyncResult) => void): void;`
         * 
         * @param coercionType - The format for the returned body.
         * @param options - Optional. An object literal that contains one or more of the following properties:
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. The body is provided in the requested format in the asyncResult.value property.
         */
        getAsync(coerciontype: Office.CoercionType, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Returns the current body in a specified format.
         *
         * This method returns the entire current body in the format specified by coercionType.
         *
         * When working with HTML-formatted bodies, it is important to note that the Body.getAsync and Body.setAsync methods are not idempotent. The value returned from the getAsync method will not necessarily be exactly the same as the value that was passed in the setAsync method previously. The client may modify the value passed to setAsync in order to make it render efficiently with its rendering engine.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         *
         * @param coercionType - The format for the returned body.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. The body is provided in the requested format in the asyncResult.value property.
         */
        getAsync(coerciontype: Office.CoercionType, callback: (result: AsyncResult) => void): void;

        /**
         * Gets a value that indicates whether the content is in HTML or text format.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr></table>
         *
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. The content type is returned as one of the CoercionType values in the asyncResult.value property.
         */
        getTypeAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Adds the specified content to the beginning of the item body.
         *
         * The prependAsync method inserts the specified string at the beginning of the item body. After insertion, the cursor is returned to its original place, relative to the inserted content.
         *
         * When including links in HTML markup, you can disable online link preview by setting the id attribute on the anchor (<a>) to "LPNoLP" (please see the Examples section for a sample).
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         * 
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         * 
         * <tr><td>Errors</td><td>DataExceedsMaximumSize - The data parameter is longer than 1,000,000 characters.</td></tr></table>
         * 
         * In addition to the main signature, this method also has these signatures:
         * 
         * `prependAsync(data: string, options: Office.AsyncContextOptions & CoercionTypeOptions): void;`
         * 
         * `prependAsync(data: string, callback: (result: AsyncResult) => void): void;`
         * 
         * `prependAsync(data: string): void;`
         *
         * @param data - The string to be inserted at the beginning of the body. The string is limited to 1,000,000 characters.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         *        coercionType: The desired format for the body. The string in the data parameter will be converted to this format.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. Any errors encountered will be provided in the asyncResult.error property.
         */
        prependAsync(data: string, options?: Office.AsyncContextOptions & CoercionTypeOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Adds the specified content to the beginning of the item body.
         *
         * The prependAsync method inserts the specified string at the beginning of the item body. After insertion, the cursor is returned to its original place, relative to the inserted content.
         *
         * When including links in HTML markup, you can disable online link preview by setting the id attribute on the anchor (<a>) to "LPNoLP" (please see the Examples section for a sample).
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         * 
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         * 
         * <tr><td>Errors</td><td>DataExceedsMaximumSize - The data parameter is longer than 1,000,000 characters.</td></tr></table>
         *
         * @param data - The string to be inserted at the beginning of the body. The string is limited to 1,000,000 characters.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         *        coercionType: The desired format for the body. The string in the data parameter will be converted to this format.
         */
        prependAsync(data: string, options: Office.AsyncContextOptions & CoercionTypeOptions): void;
        /**
         * Adds the specified content to the beginning of the item body.
         *
         * The prependAsync method inserts the specified string at the beginning of the item body. After insertion, the cursor is returned to its original place, relative to the inserted content.
         *
         * When including links in HTML markup, you can disable online link preview by setting the id attribute on the anchor (<a>) to "LPNoLP" (please see the Examples section for a sample).
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         * <tr><td>Errors</td><td>DataExceedsMaximumSize - The data parameter is longer than 1,000,000 characters.</td></tr></table>
         *
         * @param data - The string to be inserted at the beginning of the body. The string is limited to 1,000,000 characters.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. Any errors encountered will be provided in the asyncResult.error property.
         */
        prependAsync(data: string, callback: (result: AsyncResult) => void): void;
        /**
         * Adds the specified content to the beginning of the item body.
         *
         * The prependAsync method inserts the specified string at the beginning of the item body. After insertion, the cursor is returned to its original place, relative to the inserted content.
         *
         * When including links in HTML markup, you can disable online link preview by setting the id attribute on the anchor (<a>) to "LPNoLP" (please see the Examples section for a sample).
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         * <tr><td>Errors</td><td>DataExceedsMaximumSize - The data parameter is longer than 1,000,000 characters.</td></tr></table>
         *
         * @param data - The string to be inserted at the beginning of the body. The string is limited to 1,000,000 characters.
         */
        prependAsync(data: string): void;
        /**
         * Replaces the entire body with the specified text.
         *
         * When working with HTML-formatted bodies, it is important to note that the Body.getAsync and Body.setAsync methods are not idempotent. The value returned from the getAsync method will not necessarily be exactly the same as the value that was passed in the setAsync method previously. The client may modify the value passed to setAsync in order to make it render efficiently with its rendering engine.
         *
         * When including links in HTML markup, you can disable online link preview by setting the id attribute on the anchor (<a>) to "LPNoLP" (please see the Examples section for a sample).
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>DataExceedsMaximumSize - The data parameter is longer than 1,000,000 characters.</td></tr><tr><td></td><td>InvalidFormatError - The options.coercionType parameter is set to Office.CoercionType.Html and the message body is in plain text.</td></tr></table>
         * 
         * In addition to the main signature, this method also has these signatures:
         * 
         * `setAsync(data: string, options: Office.AsyncContextOptions & CoercionTypeOptions): void;`
         * 
         * `setAsync(data: string, callback: (result: AsyncResult) => void): void;`
         * 
         * `setAsync(data: string): void;`
         *
         * @param data - The string that will replace the existing body. The string is limited to 1,000,000 characters.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         *        coercionType: The desired format for the body. The string in the data parameter will be converted to this format.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.  Any errors encountered will be provided in the asyncResult.error property.
         */
        setAsync(data: string, options?: Office.AsyncContextOptions & CoercionTypeOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Replaces the entire body with the specified text.
         *
         * When working with HTML-formatted bodies, it is important to note that the Body.getAsync and Body.setAsync methods are not idempotent. The value returned from the getAsync method will not necessarily be exactly the same as the value that was passed in the setAsync method previously. The client may modify the value passed to setAsync in order to make it render efficiently with its rendering engine.
         *
         * When including links in HTML markup, you can disable online link preview by setting the id attribute on the anchor (<a>) to "LPNoLP" (please see the Examples section for a sample).
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>DataExceedsMaximumSize - The data parameter is longer than 1,000,000 characters.</td></tr><tr><td></td><td>InvalidFormatError - The options.coercionType parameter is set to Office.CoercionType.Html and the message body is in plain text.</td></tr></table>
         *
         * @param data - The string that will replace the existing body. The string is limited to 1,000,000 characters.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         *        coercionType: The desired format for the body. The string in the data parameter will be converted to this format.
         */
        setAsync(data: string, options: Office.AsyncContextOptions & CoercionTypeOptions): void;
        /**
         * Replaces the entire body with the specified text.
         *
         * When working with HTML-formatted bodies, it is important to note that the Body.getAsync and Body.setAsync methods are not idempotent. The value returned from the getAsync method will not necessarily be exactly the same as the value that was passed in the setAsync method previously. The client may modify the value passed to setAsync in order to make it render efficiently with its rendering engine.
         *
         * When including links in HTML markup, you can disable online link preview by setting the id attribute on the anchor (<a>) to "LPNoLP" (please see the Examples section for a sample).
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>DataExceedsMaximumSize - The data parameter is longer than 1,000,000 characters.</td></tr><tr><td></td><td>InvalidFormatError - The options.coercionType parameter is set to Office.CoercionType.Html and the message body is in plain text.</td></tr></table>
         *
         * @param data - The string that will replace the existing body. The string is limited to 1,000,000 characters.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.  Any errors encountered will be provided in the asyncResult.error property.
         */
        setAsync(data: string, callback: (result: AsyncResult) => void): void;
        /**
         * Replaces the entire body with the specified text.
         *
         * When working with HTML-formatted bodies, it is important to note that the Body.getAsync and Body.setAsync methods are not idempotent. The value returned from the getAsync method will not necessarily be exactly the same as the value that was passed in the setAsync method previously. The client may modify the value passed to setAsync in order to make it render efficiently with its rendering engine.
         *
         * When including links in HTML markup, you can disable online link preview by setting the id attribute on the anchor (<a>) to "LPNoLP" (please see the Examples section for a sample).
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>DataExceedsMaximumSize - The data parameter is longer than 1,000,000 characters.</td></tr><tr><td></td><td>InvalidFormatError - The options.coercionType parameter is set to Office.CoercionType.Html and the message body is in plain text.</td></tr></table>
         *
         * @param data - The string that will replace the existing body. The string is limited to 1,000,000 characters.
         */
        setAsync(data: string): void;

        /**
         * Replaces the selection in the body with the specified text.
         *
         * The setSelectedDataAsync method inserts the specified string at the cursor location in the body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor was never in the body of the item, or if the body of the item lost focus in the UI, the string will be inserted at the top of the body content. After insertion, the cursor is placed at the end of the inserted content.
         *
         * When including links in HTML markup, you can disable online link preview by setting the id attribute on the anchor (<a>) to "LPNoLP" (please see the Examples section for a sample).
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>DataExceedsMaximumSize - The data parameter is longer than 1,000,000 characters.</td></tr><tr><td></td><td>InvalidFormatError - The options.coercionType parameter is set to Office.CoercionType.Html and the message body is in plain text.</td></tr></table>
         * 
         * In addition to the main signature, this method also has these signatures:
         * 
         * `setSelectedDataAsync(data: string, options: Office.AsyncContextOptions & CoercionTypeOptions): void;`
         * 
         * `setSelectedDataAsync(data: string, callback: (result: AsyncResult) => void): void;`
         * 
         * `setSelectedDataAsync(data: string): void;`
         *         
         * @param data - The string that will replace the existing body. The string is limited to 1,000,000 characters.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         *        coercionType: The desired format for the body. The string in the data parameter will be converted to this format.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.  Any errors encountered will be provided in the asyncResult.error property.
         */
        setSelectedDataAsync(data: string, options?: Office.AsyncContextOptions & CoercionTypeOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Replaces the selection in the body with the specified text.
         *
         * The setSelectedDataAsync method inserts the specified string at the cursor location in the body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor was never in the body of the item, or if the body of the item lost focus in the UI, the string will be inserted at the top of the body content. After insertion, the cursor is placed at the end of the inserted content.
         *
         * When including links in HTML markup, you can disable online link preview by setting the id attribute on the anchor (<a>) to "LPNoLP" (please see the Examples section for a sample).
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>DataExceedsMaximumSize - The data parameter is longer than 1,000,000 characters.</td></tr><tr><td></td><td>InvalidFormatError - The options.coercionType parameter is set to Office.CoercionType.Html and the message body is in plain text.</td></tr></table>
         *
         * @param data - The string that will replace the existing body. The string is limited to 1,000,000 characters.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         *        coercionType: The desired format for the body. The string in the data parameter will be converted to this format.
         */
        setSelectedDataAsync(data: string, options: Office.AsyncContextOptions & CoercionTypeOptions): void;
        /**
         * Replaces the selection in the body with the specified text.
         *
         * The setSelectedDataAsync method inserts the specified string at the cursor location in the body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor was never in the body of the item, or if the body of the item lost focus in the UI, the string will be inserted at the top of the body content. After insertion, the cursor is placed at the end of the inserted content.
         *
         * When including links in HTML markup, you can disable online link preview by setting the id attribute on the anchor (<a>) to "LPNoLP" (please see the Examples section for a sample).
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>DataExceedsMaximumSize - The data parameter is longer than 1,000,000 characters.</td></tr><tr><td></td><td>InvalidFormatError - The options.coercionType parameter is set to Office.CoercionType.Html and the message body is in plain text.</td></tr></table>
         *
         * @param data - The string that will replace the existing body. The string is limited to 1,000,000 characters.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.  Any errors encountered will be provided in the asyncResult.error property.
         */
        setSelectedDataAsync(data: string, callback: (result: AsyncResult) => void): void;
        /**
         * Replaces the selection in the body with the specified text.
         *
         * The setSelectedDataAsync method inserts the specified string at the cursor location in the body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor was never in the body of the item, or if the body of the item lost focus in the UI, the string will be inserted at the top of the body content. After insertion, the cursor is placed at the end of the inserted content.
         *
         * When including links in HTML markup, you can disable online link preview by setting the id attribute on the anchor (<a>) to "LPNoLP" (please see the Examples section for a sample).
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>DataExceedsMaximumSize - The data parameter is longer than 1,000,000 characters.</td></tr><tr><td></td><td>InvalidFormatError - The options.coercionType parameter is set to Office.CoercionType.Html and the message body is in plain text.</td></tr></table>
         *
         * @param data - The string that will replace the existing body. The string is limited to 1,000,000 characters.
         */
        setSelectedDataAsync(data: string): void;
    }
    /**
     * Represents a contact stored on the server. Read mode only.
     *
     * The list of contacts associated with an email message or appointment is returned in the contacts property of the {@link Office.Entities} object that is returned by the getEntities or getEntitiesByType method of the active item.
     *
     * [Api set: Mailbox 1.0]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>Restricted</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
     */
    export interface Contact {
        /**
         * An array of strings containing the mailing and street addresses associated with the contact. Nullable.
         */
        addresses: Array<string>;
        /**
         * A string containing the name of the business associated with the contact. Nullable.
         */
        businessName: string;
        /**
         * An array of strings containing the SMTP email addresses associated with the contact. Nullable,
         */
        emailAddresses: Array<string>;
        /**
         * A string containing the name of the person associated with the contact. Nullable.
         */
        personName: string;
        /**
         * An array containing a PhoneNumber object for each phone number associated with the contact. Nullable.
         */
        phoneNumbers: Array<PhoneNumber>;
        /**
         * An array of strings containing the Internet URLs associated with the contact. Nullable.
         */
        urls: Array<string>;
    }
    /**
     * The CustomProperties object represents custom properties that are specific to a particular item and specific to a mail add-in for Outlook. For example, there might be a need for a mail add-in to save some data that is specific to the current email message that activated the add-in. If the user revisits the same message in the future and activates the mail add-in again, the add-in will be able to retrieve the data that had been saved as custom properties.
     *
     * Because Outlook for Mac doesn't cache custom properties, if the user's network goes down, mail add-ins cannot access their custom properties.
     *
     * [Api set: Mailbox 1.0]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
     */
    export interface CustomProperties {
        /**
         * Returns the value of the specified custom property.
         * @param name - The name of the custom property to be returned.
         * @returns The value of the specified custom property.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        get(name: string): any;
        /**
         * Sets the specified property to the specified value.
         *
         * The set method sets the specified property to the specified value. You must use the saveAsync method to save the property to the server.
         *
         * The set method creates a new property if the specified property does not already exist; otherwise, the existing value is replaced with the new value. The value parameter can be of any type; however, it is always passed to the server as a string.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         *
         * @param name - The name of the property to be set.
         * @param value - The value of the property to be set.
         */
        set(name: string, value: string): void;
        /**
         * Removes the specified property from the custom property collection.
         *
         * To make the removal of the property permanent, you must call the saveAsync method of the CustomProperties object.
         * @param name - The name of the property to be removed.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        remove(name: string): void;
        /**
         * Saves item-specific custom properties to the server.
         *
         * You must call the saveAsync method to persist any changes made with the set method or the remove method of the CustomProperties object. The saving action is asynchronous.
         *
         * It's a good practice to have your callback function check for and handle errors from saveAsync. In particular, a read add-in can be activated while the user is in a connected state in a read form, and subsequently the user becomes disconnected. If the add-in calls saveAsync while in the disconnected state, saveAsync would return an error. Your callback method should handle this error accordingly.
         *
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         * @param asyncContext - Optional. Any state data that is passed to the callback method.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        saveAsync(callback?: (result: AsyncResult) => void, asyncContext?: any): void;
    }
    /**
     * Provides diagnostic information to an Outlook add-in.
     *
     * [Api set: Mailbox 1.0]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
     */
    export interface Diagnostics {
        /**
         * Gets a string that represents the name of the host application.
         *
         * A string that can be one of the following values: Outlook, Mac Outlook, OutlookIOS, or OutlookWebApp.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        hostName: string;
        /**
         * Gets a string that represents the version of either the host application or the Exchange Server.
         *
         * If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the hostVersion property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string 15.0.468.0.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        hostVersion: string;
        /**
         * Gets a string that represents the current view of Outlook Web App.
         *
         * The returned string can be one of the following values: OneColumn, TwoColumns, or ThreeColumns.
         *
         * If the host application is not Outlook Web App, then accessing this property results in undefined.
         *
         * Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:
         *
         * - OneColumn, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.
         *
         * - TwoColumns, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.
         *
         * - ThreeColumns, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        OWAView: "string";
    }
    /**
     * Provides the email properties of the sender or specified recipients of an email message or appointment.
     *
     * [Api set: Mailbox 1.0]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
     */
    export interface EmailAddressDetails {
        /**
         * Gets the SMTP email address.
         */
        emailAddress: string;
        /**
         * Gets the display name associated with an email address.
         */
        displayName: string;
        /**
         * Gets the response that an attendee returned for an appointment. This property applies to only an attendee of an appointment, as represented by the optionalAttendees or requiredAttendees property. This property returns undefined in other scenarios.
         */
        appointmentResponse: Office.MailboxEnums.ResponseType;
        /**
         * Gets the email address type of a recipient.
         */
        recipientType: Office.MailboxEnums.RecipientType;
    }
    /**
     * Represents an email account on an Exchange Server.
     *
     * [Api set: Mailbox 1.0]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
     */
    export interface EmailUser {
        /**
         * Gets the display name associated with an email address.
         */
        displayName: string;
        /**
         * Gets the SMTP email address.
         */
        emailAddress: string;
    }
    /**
     * Represents a collection of entities found in an email message or appointment. Read mode only.
     *
     * The Entities object is a container for the entity arrays returned by the getEntities and getEntitiesByType methods when the item (either an email message or an appointment) contains one or more entities that have been found by the server. You can use these entities in your code to provide additional context information to the viewer, such as a map to an address found in the item, or to open a dialer for a phone number found in the item.
     *
     * If no entities of the type specified in the property are present in the item, the property associated with that entity is null. For example, if a message contains a street address and a phone number, the addresses property and phoneNumbers property would contain information, and the other properties would be null.
     *
     * To be recognized as an address, the string must contain a United States postal address that has at least a subset of the elements of a street number, street name, city, state, and zip code.
     *
     * To be recognized as a phone number, the string must contain a North American phone number format.
     *
     * Entity recognition relies on natural language recognition that is based on machine learning of large amounts of data. The recognition of an entity is non-deterministic and success sometimes relies on the particular context in the item.
     *
     * When the property arrays are returned by the getEntitiesByType method, only the property for the specified entity contains data; all other properties are null.
     *
     * [Api set: Mailbox 1.0]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
     */
    export interface Entities {
        /**
         * Gets the physical addresses (street or mailing addresses) found in an email message or appointment.
         */
        addresses: string[];
        /**
         * Gets the contacts found in an email address or appointment.
         */
        contacts: Contact[];
        /**
         * Gets the email addresses found in an email message or appointment.
         */
        emailAddresses: string[];
        /**
         * Gets the meeting suggestions found in an email message.
         */
        meetingSuggestions: MeetingSuggestion[];
        /**
         * Gets the phone numbers found in an email message or appointment.
         */
        phoneNumbers: PhoneNumber[];
        /**
         * Gets the task suggestions found in an email message or appointment.
         */
        taskSuggestions: string[];
        /**
         * Gets the Internet URLs present in an email message or appointment.
         */
        urls: string[];
    }
    /**
     * The subclass of {@link Office.Item} dealing with appointments.
     * 
     * Important: This is an internal Outlook object, not directly exposed through existing interfaces. You should treat this as a mode of Office.context.mailbox.item. Refer to the Object Model pages for more information.
     */
    export interface Appointment extends Item {
    }
    /**
     * The appointment organizer mode of {@link Office.Item | Office.context.mailbox.item}.
     * 
     * Important: This is an internal Outlook object, not directly exposed through existing interfaces. You should treat this as a mode of `Office.context.mailbox.item`. Refer to the Object Model pages for more information.
     */
    export interface AppointmentCompose extends Appointment, ItemCompose {
         /**
         * Gets an object that provides methods for manipulating the body of an item.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr></table>
         */
        body: Office.Body;
        /**
         * Gets the date and time that an item was created.  Read mode only.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr></table>
         */
        dateTimeCreated: Date;
        /**
         * Gets the date and time that an item was last modified. Read mode only.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr></table>
         *
         * Note: This member is not supported in Outlook for iOS or Outlook for Android.
         */
        dateTimeModifed: Date;
        /**
         * Gets or sets the date and time that the appointment is to end.
         *
         * The end property is an {@link Office.Time} object expressed as a Coordinated Universal Time (UTC) date and time value. You can use the convertToLocalClientTime method to convert the end property value to the client's local date and time.
         *
         * When you use the Time.setAsync method to set the end time, you should use the convertToUtcClientTime method to convert the local time on the client to UTC for the server.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr></table>
         */
        end: Time;
        /**
         * Gets the type of item that an instance represents.
         *
         * The itemType property returns one of the ItemType enumeration values, indicating whether the item object instance is a message or an appointment.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr></table>
         */
        itemType: Office.MailboxEnums.ItemType;
        /**
         * Gets or sets the {@link Office.Location} of an appointment. The location property returns a Location object that provides methods that are used to get and set the location of the appointment.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr></table>
         */
        location: Location;
        /**
         * Gets the notification messages for an item.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr></table>
         */
        notificationMessages: Office.NotificationMessages;
        /**
         * Provides access to the optional attendees of an event. The type of object and level of access depends on the mode of the current item. The optionalAttendees property returns an {@link Office.Recipients} object that provides methods to get or update the optional attendees for a meeting.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr></table>
         */
        optionalAttendees: Recipients;
        /**
         * Provides access to the required attendees of an event. The type of object and level of access depends on the mode of the current item. The requiredAttendees property returns an {@link Office.Recipients} object that provides methods to get or update the required attendees for a meeting.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr></table>
         */
        requiredAttendees: Recipients;
        /**
         * Gets or sets the date and time that the appointment is to begin.
         *
         * The start property is an {@link Office.Time} object expressed as a Coordinated Universal Time (UTC) date and time value. You can use the convertToLocalClientTime method to convert the value to the client's local date and time.
         *
         * When you use the Time.setAsync method to set the start time, you should use the convertToUtcClientTime method to convert the local time on the client to UTC for the server.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr></table>
         */
        start: Time;
        /**
         * Gets or sets the description that appears in the subject field of an item.
         *
         * The subject property gets or sets the entire subject of the item, as sent by the email server.
         *
         * The subject property returns a Subject object that provides methods to get and set the subject.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr></table>
         */
        subject: Subject;
        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The addFileAttachmentAsync method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr>
         *
         * <tr><td>Errors</td><td>AttachmentSizeExceeded - The attachment is larger than allowed.</td></tr><tr><td></td><td>FileTypeNotSupported - The attachment has an extension that is not allowed.</td></tr><tr><td></td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         * 
         * In addition to this signature, the method also has the following signatures:
         * 
         * `addFileAttachmentAsync(uri: string, attachmentName: string): void;`
         * 
         * `addFileAttachmentAsync(uri: string, attachmentName: string, options: Office.AsyncContextOptions): void;`
         * 
         * `addFileAttachmentAsync(uri: string, attachmentName: string, callback: (result: AsyncResult) => void): void;`
         *
         * @param uri - The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         *        isInline: If true, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type asyncResult. On success, the attachment identifier will be provided in the asyncResult.value property. If uploading the attachment fails, the asyncResult object will contain an Error object that provides a description of the error.
         */
        addFileAttachmentAsync(uri: string, attachmentName: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The addFileAttachmentAsync method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr>
         *
         * <tr><td>Errors</td><td>AttachmentSizeExceeded - The attachment is larger than allowed.</td></tr><tr><td></td><td>FileTypeNotSupported - The attachment has an extension that is not allowed.</td></tr><tr><td></td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         *
         * @param uri - The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         */
        addFileAttachmentAsync(uri: string, attachmentName: string): void;
        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The addFileAttachmentAsync method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr>
         *
         * <tr><td>Errors</td><td>AttachmentSizeExceeded - The attachment is larger than allowed.</td></tr><tr><td></td><td>FileTypeNotSupported - The attachment has an extension that is not allowed.</td></tr><tr><td></td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         *
         * @param uri - The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         *        isInline: If true, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.
         */
        addFileAttachmentAsync(uri: string, attachmentName: string, options: Office.AsyncContextOptions): void;
        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The addFileAttachmentAsync method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr>
         *
         * <tr><td>Errors</td><td>AttachmentSizeExceeded - The attachment is larger than allowed.</td></tr><tr><td></td><td>FileTypeNotSupported - The attachment has an extension that is not allowed.</td></tr><tr><td></td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         *
         * @param uri - The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type asyncResult. On success, the attachment identifier will be provided in the asyncResult.value property. If uploading the attachment fails, the asyncResult object will contain an Error object that provides a description of the error.
         */
        addFileAttachmentAsync(uri: string, attachmentName: string, callback: (result: AsyncResult) => void): void;
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message or appointment.
         *
         * The addItemAttachmentAsync method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, asyncResult, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the options parameter to pass state information to the callback method, if needed.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * If your Office Add-in is running in Outlook Web App, the addItemAttachmentAsync method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr>
         *
         * <tr><td>Errors</td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         * 
         * In addition to the main signature, this method also has these signatures:
         * 
         * `addItemAttachmentAsync(itemId: any, attachmentName: string): void;`
         * 
         * `addItemAttachmentAsync(itemId: any, attachmentName: string, options: Office.AsyncContextOptions): void;`
         * 
         * `addItemAttachmentAsync(itemId: any, attachmentName: string, callback: (result: AsyncResult) => void): void;`
         *
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. On success, the attachment identifier will be provided in the asyncResult.value property. If adding the attachment fails, the asyncResult object will contain an Error object that provides a description of the error.
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message or appointment.
         *
         * The addItemAttachmentAsync method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, asyncResult, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the options parameter to pass state information to the callback method, if needed.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * If your Office Add-in is running in Outlook Web App, the addItemAttachmentAsync method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr>
         *
         * <tr><td>Errors</td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         *
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string): void;
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message or appointment.
         *
         * The addItemAttachmentAsync method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, asyncResult, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the options parameter to pass state information to the callback method, if needed.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * If your Office Add-in is running in Outlook Web App, the addItemAttachmentAsync method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr>
         *
         * <tr><td>Errors</td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         *
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string, options: Office.AsyncContextOptions): void;
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message or appointment.
         *
         * The addItemAttachmentAsync method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, asyncResult, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the options parameter to pass state information to the callback method, if needed.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * If your Office Add-in is running in Outlook Web App, the addItemAttachmentAsync method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr>
         *
         * <tr><td>Errors</td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         *
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. On success, the attachment identifier will be provided in the asyncResult.value property. If adding the attachment fails, the asyncResult object will contain an Error object that provides a description of the error.
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string, callback: (result: AsyncResult) => void): void;
        /**
         * Closes the current item that is being composed
         *
         * The behaviors of the close method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or close the action.
         *
         * In the Outlook desktop client, if the message is an inline reply, the close method has no effect.
         *
         * Note: In Outlook on the web, if the item is an appointment and it has previously been saved using saveAsync, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>Restricted</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr></table>
         */
        close(): void;
        /**
         * Asynchronously returns selected data from the subject or body of a message.
         *
         * If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the InvalidSelection error.
         *
         * To access the selected data from the callback method, call asyncResult.value.data. To access the source property that the selection comes from, call asyncResult.value.sourceProperty, which will be either body or subject.
         *
         * [Api set: Mailbox 1.2]
         *
         * @returns
         * The selected data as a string with format determined by coercionType.
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr></table>
         *
         * @param coercionType - Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.
         * @param options - An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         */
        getSelectedDataAsync(coerciontype: Office.CoercionType, options: Office.AsyncContextOptions, callback: (result: AsyncResult) => void): void;
         /**
         * Asynchronously returns selected data from the subject or body of a message.
         *
         * If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the InvalidSelection error.
         *
         * To access the selected data from the callback method, call asyncResult.value.data. To access the source property that the selection comes from, call asyncResult.value.sourceProperty, which will be either body or subject.
         *
         * [Api set: Mailbox 1.2]
         *
         * @returns
         * The selected data as a string with format determined by coercionType.
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr></table>
         *
         * @param coercionType - Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         */
        getSelectedDataAsync(coerciontype: Office.CoercionType, callback: (result: AsyncResult) => void): void;
        /**
         * Asynchronously loads custom properties for this add-in on the selected item.
         *
         * Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a CustomProperties object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.
         *
         * The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr></table>
         *
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         * @param userContext - Optional. Developers can provide any object they wish to access in the callback function. This object can be accessed by the asyncResult.asyncContext property in the callback function.
         */
        loadCustomPropertiesAsync(callback: (result: AsyncResult) => void, userContext?: any): void;
        /**
         * Removes an attachment from a message or appointment.
         *
         * The removeAttachmentAsync method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         * 
         * In addition to this signature, the method also has the following signatures:
         * 
         * `removeAttachmentAsync(attachmentIndex: string): void;`
         * 
         * `removeAttachmentAsync(attachmentIndex: string, options: Office.AsyncContextOptions): void;`
         * 
         * `removeAttachmentAsync(attachmentIndex: string, callback: (result: AsyncResult) => void): void;`
         *
         * @param attachmentIndex - The identifier of the attachment to remove. The maximum length of the string is 100 characters.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If removing the attachment fails, the asyncResult.error property will contain an error code with the reason for the failure.
         */
        removeAttachmentAsync(attachmentIndex: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Removes an attachment from a message or appointment.
         *
         * The removeAttachmentAsync method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param attachmentIndex - The identifier of the attachment to remove. The maximum length of the string is 100 characters.
         */
        removeAttachmentAsync(attachmentIndex: string): void;
        /**
         * Removes an attachment from a message or appointment.
         *
         * The removeAttachmentAsync method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param attachmentIndex - The identifier of the attachment to remove. The maximum length of the string is 100 characters.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         */
        removeAttachmentAsync(attachmentIndex: string, options: Office.AsyncContextOptions): void;
        /**
         * Removes an attachment from a message or appointment.
         *
         * The removeAttachmentAsync method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param attachmentIndex - The identifier of the attachment to remove. The maximum length of the string is 100 characters.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If removing the attachment fails, the asyncResult.error property will contain an error code with the reason for the failure.
         */
        removeAttachmentAsync(attachmentIndex: string, callback: (result: AsyncResult) => void): void;
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.
         *
         * Since appointments have no draft state, if saveAsync is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.
         *
         * Note: If your add-in calls saveAsync on an item in compose mode in order to get an itemId to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the itemId will return an error.
         *
         * Note: The following clients have different behavior for saveAsync on appointments in compose mode:
         *
         * - Mac Outlook does not support saveAsync on a meeting in compose mode. Calling saveAsync on a meeting in Mac Outlook will return an error.
         *
         * - Outlook on the web always sends an invitation or update when saveAsync is called on an appointment in compose mode.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         * 
         * In addition to this signature, the method also has the following signatures:
         * 
         * `saveAsync(): void;`
         * 
         * `saveAsync(options: Office.AsyncContextOptions): void;`
         * 
         * `saveAsync(callback: (result: AsyncResult) => void): void;`
         *
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If removing the attachment fails, the asyncResult.error property will contain an error code with the reason for the failure.
         */
        saveAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.
         *
         * Since appointments have no draft state, if saveAsync is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.
         *
         * Note: If your add-in calls saveAsync on an item in compose mode in order to get an itemId to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the itemId will return an error.
         *
         * Note: The following clients have different behavior for saveAsync on appointments in compose mode:
         *
         * - Mac Outlook does not support saveAsync on a meeting in compose mode. Calling saveAsync on a meeting in Mac Outlook will return an error.
         *
         * - Outlook on the web always sends an invitation or update when saveAsync is called on an appointment in compose mode.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         */
        saveAsync(): void;
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.
         *
         * Since appointments have no draft state, if saveAsync is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.
         *
         * Note: If your add-in calls saveAsync on an item in compose mode in order to get an itemId to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the itemId will return an error.
         *
         * Note: The following clients have different behavior for saveAsync on appointments in compose mode:
         *
         * - Mac Outlook does not support saveAsync on a meeting in compose mode. Calling saveAsync on a meeting in Mac Outlook will return an error.
         *
         * - Outlook on the web always sends an invitation or update when saveAsync is called on an appointment in compose mode.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         */
        saveAsync(options: Office.AsyncContextOptions): void;
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.
         *
         * Since appointments have no draft state, if saveAsync is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.
         *
         * Note: If your add-in calls saveAsync on an item in compose mode in order to get an itemId to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the itemId will return an error.
         *
         * Note: The following clients have different behavior for saveAsync on appointments in compose mode:
         *
         * - Mac Outlook does not support saveAsync on a meeting in compose mode. Calling saveAsync on a meeting in Mac Outlook will return an error.
         *
         * - Outlook on the web always sends an invitation or update when saveAsync is called on an appointment in compose mode.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If removing the attachment fails, the asyncResult.error property will contain an error code with the reason for the failure.
         */
        saveAsync(callback: (result: AsyncResult) => void): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         *
         * The setSelectedDataAsync method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.
         *
         * [Api set: Mailbox 1.2]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         * 
         * In addition to this signature, the method also has the following signatures:
         * 
         * `setSelectedDataAsync(data: string): void;`
         * 
         * `setSelectedDataAsync(data: string, options: Office.AsyncContextOptions & CoercionTypeOptions): void;`
         * 
         * `setSelectedDataAsync(data: string, callback: (result: AsyncResult) => void): void;`
         *
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an ArgumentOutOfRange exception is thrown.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         *        coercionType: If text, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML. If html and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an InvalidDataFormat error is returned. If coercionType is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If removing the attachment fails, the asyncResult.error property will contain an error code with the reason for the failure.
         */
        setSelectedDataAsync(data: string, options?: Office.AsyncContextOptions & CoercionTypeOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         *
         * The setSelectedDataAsync method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.
         *
         * [Api set: Mailbox 1.2]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an ArgumentOutOfRange exception is thrown.
         */
        setSelectedDataAsync(data: string): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         *
         * The setSelectedDataAsync method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.
         *
         * [Api set: Mailbox 1.2]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an ArgumentOutOfRange exception is thrown.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         *        coercionType: If text, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML. If html and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an InvalidDataFormat error is returned. If coercionType is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.
         */
        setSelectedDataAsync(data: string, options: Office.AsyncContextOptions & CoercionTypeOptions): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         *
         * The setSelectedDataAsync method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.
         *
         * [Api set: Mailbox 1.2]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Organizer</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an ArgumentOutOfRange exception is thrown.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If removing the attachment fails, the asyncResult.error property will contain an error code with the reason for the failure.
         */
        setSelectedDataAsync(data: string, callback: (result: AsyncResult) => void): void;
    }

    /**
     * The appointment attendee mode of {@link Office.Item | Office.context.mailbox.item}.
     * 
     * Important: This is an internal Outlook object, not directly exposed through existing interfaces. You should treat this as a mode of 'Office.context.mailbox.item'. Refer to the Object Model pages for more information.
     */
    export interface AppointmentRead extends Appointment, ItemRead {
        /**
         * Gets an array of attachments for the item.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * 
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         *
         * Note: Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned. For more information, see {@link https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519 | Blocked attachments in Outlook}.
         *
         */
        attachments: Office.AttachmentDetails[];
        /**
         * Gets an object that provides methods for manipulating the body of an item.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         */
        body: Office.Body;
        /**
         * Gets the date and time that an item was created. Read mode only.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         */
        dateTimeCreated: Date;
        /**
         * Gets the date and time that an item was last modified. Read mode only.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         *
         * Note: This member is not supported in Outlook for iOS or Outlook for Android.
         */
        dateTimeModifed: Date;
        /**
         * Gets the date and time that the appointment is to end.
         *
         * The end property is a Date object expressed as a Coordinated Universal Time (UTC) date and time value. You can use the convertToLocalClientTime method to convert the end property value to the client's local date and time.
         *
         * When you use the Time.setAsync method to set the end time, you should use the convertToUtcClientTime method to convert the local time on the client to UTC for the server.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         */
        end: Date;
        /**
         * Gets the Exchange Web Services item class of the selected item.
         *
         *
         * You can create custom message classes that extends a default message class, for example, a custom appointment message class IPM.Appointment.Contoso.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         * 
         * The itemClass property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.
         * 
         * <table>
         *   <tr>
         *     <th>Type</th>
         *     <th>Description</th>
         *     <th>Item Class</th>
         *   </tr>
         *   <tr>
         *     <td>Appointment items</td>
         *     <td>These are calendar items of the item class IPM.Appointment or IPM.Appointment.Occurence.</td>
         *     <td>IPM.Appointment,IPM.Appointment.Occurence</td>
         *   </tr>
         *   <tr>
         *     <td>Message items</td>
         *     <td>These include email messages that have the default message class IPM.Note, and meeting requests, responses, and cancellations, that use IPM.Schedule.Meeting as the base message class.</td>
         *     <td>IPM.Note,IPM.Schedule.Meeting.Request,IPM.Schedule.Meeting.Neg,IPM.Schedule.Meeting.Pos,IPM.Schedule.Meeting.Tent,IPM.Schedule.Meeting.Canceled</td>
         *   </tr>
         * </table>
         * 
         */
        itemClass: string;
        /**
         * Gets the Exchange Web Services item identifier for the current item.
         *
         * The itemId property is not available in compose mode. If an item identifier is required, the saveAsync method can be used to save the item to the store, which will return the item identifier in the AsyncResult.value parameter in the callback function.
         *
         * Note: The identifier returned by the itemId property is the same as the Exchange Web Services item identifier. The itemId property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API. Before making REST API calls using this value, it should be converted using Office.context.mailbox.convertToRestId. For more details, see Use the Outlook REST APIs from an Outlook add-in.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         */
        itemId: string;
        /**
         * Gets the type of item that an instance represents.
         *
         * The itemType property returns one of the ItemType enumeration values, indicating whether the item object instance is a message or an appointment.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         */
        itemType: Office.MailboxEnums.ItemType;
        /**
         * Gets the location of an appointment.
         *
         * The location property returns a string that contains the location of the appointment.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         */
        location: string;
        /**
         * Gets the subject of an item, with all prefixes removed (including RE: and FWD:).
         *
         * The normalizedSubject property gets the subject of the item, with any standard prefixes (such as RE: and FW:) that are added by email programs. To get the subject of the item with the prefixes intact, use the subject property.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         */
        normalizedSubject: string;
        /**
         * Gets the notification messages for an item.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         */
        notificationMessages: Office.NotificationMessages;
        /**
         * Provides access to the optional attendees of an event. The type of object and level of access depends on the mode of the current item.
         *
         * The optionalAttendees property returns an array that contains an {@link Office.EmailAddressDetails} object for each optional attendee to the meeting.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         */
        optionalAttendees: EmailAddressDetails[];
        /**
         * Gets the email address of the meeting organizer for a specified meeting. Read mode only.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         */
        organizer: EmailAddressDetails;
        /**
         * Provides access to the required attendees of an event. The type of object and level of access depends on the mode of the current item.
         *
         * The requiredAttendees property returns an array that contains an {@link Office.EmailAddressDetails} object for each required attendee to the meeting.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         */
        requiredAttendees: EmailAddressDetails[];
        /**
         * Gets the date and time that the appointment is to begin.
         *
         * The start property is a Date object expressed as a Coordinated Universal Time (UTC) date and time value. You can use the convertToLocalClientTime method to convert the value to the client's local date and time.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         */
        start: Date;
        /**
         * Gets the description that appears in the subject field of an item.
         *
         * The subject property gets or sets the entire subject of the item, as sent by the email server.
         *
         * The subject property returns a string. Use the normalizedSubject property to get the subject minus any leading prefixes such as RE: and FW:.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         */
        subject: string;
        /**
         * Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.
         *
         * In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.
         *
         * If any of the string parameters exceed their limits, displayReplyAllForm throws an exception.
         *
         * When attachments are specified in the formData.attachments parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *  OR
         * An {@link Office.ReplyFormData} object that contains body or attachment data and a callback function
         */
        displayReplyAllForm(formData: string | ReplyFormData): void;
        /**
         * Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.
         *
         * In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.
         *
         * If any of the string parameters exceed their limits, displayReplyForm throws an exception.
         *
         * When attachments are specified in the formData.attachments parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.
         * OR
         * An {@link Office.ReplyFormData} object that contains body or attachment data and a callback function.
         */
        displayReplyForm(formData: string | ReplyFormData): void;
        /**
         * Gets the entities found in the selected item.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         */
        getEntities(): Entities;
        /**
         * Gets an array of all the entities of the specified entity type found in the selected item.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         * 
         * @param entityType - One of the EntityType enumeration values.
         *
         * @returns
         * If the value passed in entityType is not a valid member of the EntityType enumeration, the method returns null. If no entities of the specified type are present on the item, the method returns an empty array. Otherwise, the type of the objects in the returned array depends on the type of entity requested in the entityType parameter.
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>Restricted</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         * 
         * While the minimum permission level to use this method is Restricted, some entity types require ReadItem to access, as specified in the following table.
         * 
         * <table>
         *   <tr>
         *     <th>Value of entityType</th>
         *     <th>Type of objects in returned array</th>
         *     <th>Required Permission Leve</th>
         *   </tr>
         *   <tr>
         *     <td>Address</td>
         *     <td>String</td>
         *     <td>Restricted</td>
         *   </tr>
         *   <tr>
         *     <td>Contact</td>
         *     <td>Contact</td>
         *     <td>ReadItem</td>
         *   </tr>
         *   <tr>
         *     <td>EmailAddress</td>
         *     <td>String</td>
         *     <td>ReadItem</td>
         *   </tr>
         *   <tr>
         *     <td>MeetingSuggestion</td>
         *     <td>MeetingSuggestion</td>
         *     <td>ReadItem</td>
         *   </tr>
         *   <tr>
         *     <td>PhoneNumber</td>
         *     <td>PhoneNumber</td>
         *     <td>Restricted</td>
         *   </tr>
         *   <tr>
         *     <td>TaskSuggestion</td>
         *     <td>TaskSuggestion</td>
         *     <td>ReadItem</td>
         *   </tr>
         *   <tr>
         *     <td>URL</td>
         *     <td>String</td>
         *     <td>Restricted</td>
         *   </tr>
         * </table>
         */
        getEntitiesByType(entityType: Office.MailboxEnums.EntityType): (string | Contact | MeetingSuggestion | PhoneNumber | TaskSuggestion)[];
        /**
         * Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.
         *
         * The getFilteredEntitiesByName method returns the entities that match the regular expression defined in the ItemHasKnownEntity rule element in the manifest XML file with the specified FilterName element value.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         *
         * @param name - The name of the ItemHasKnownEntity rule element that defines the filter to match.
         * @returns If there is no ItemHasKnownEntity element in the manifest with a FilterName element value that matches the name parameter, the method returns null. If the name parameter does match an ItemHasKnownEntity element in the manifest, but there are no entities in the current item that match, the method return an empty array.
         */
        getFilteredEntitiesByName(name: string): (string | Contact | MeetingSuggestion | PhoneNumber | TaskSuggestion)[];
        /**
         * Returns string values in the selected item that match the regular expressions defined in the manifest XML file.
         *
         * The getRegExMatches method returns the strings that match the regular expression defined in each ItemHasRegularExpressionMatch or ItemHasKnownEntity rule element in the manifest XML file. For an ItemHasRegularExpressionMatch rule, a matching string has to occur in the property of the item that is specified by that rule. The PropertyName simple type defines the supported properties.
         *
         * If you specify an ItemHasRegularExpressionMatch rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as .* to obtain the entire body of an item does not always return the expected results. Instead, use the Body.getAsync method to retrieve the entire body.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @returns
         * An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the RegExName attribute of the matching ItemHasRegularExpressionMatch rule or the FilterName attribute of the matching ItemHasKnownEntity rule.
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         */
        getRegExMatches(): any;
        /**
         * Returns string values in the selected item that match the named regular expression defined in the manifest XML file.
         *
         * The getRegExMatchesByName method returns the strings that match the regular expression defined in the ItemHasRegularExpressionMatch rule element in the manifest XML file with the specified RegExName element value.
         *
         * If you specify an ItemHasRegularExpressionMatch rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as .* to obtain the entire body of an item does not always return the expected results.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @returns
         * An array that contains the strings that match the regular expression defined in the manifest XML file.
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
         *
         * @param name - The name of the ItemHasRegularExpressionMatch rule element that defines the filter to match.
         */
        getRegExMatchesByName(name: string): string[];
       /**
        * Asynchronously loads custom properties for this add-in on the selected item.
        *
        * Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a CustomProperties object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.
        *
        * The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.
        *
        * [Api set: Mailbox 1.0]
        *
        * @remarks
        *
        * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
        *
        * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Appointment Attendee</td></tr></table>
        *
        * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
        * @param userContext - Optional. Developers can provide any object they wish to access in the callback function. This object can be accessed by the asyncResult.asyncContext property in the callback function.
        */
       loadCustomPropertiesAsync(callback: (result: AsyncResult) => void, userContext?: any): void;
    }

    /**
     * The item namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the item by using the `itemType` property.
     *
     * [Api set: Mailbox 1.0]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>Restricted</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
     */
    export interface Item {
        /**
         * Gets an object that provides methods for manipulating the body of an item.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        body: Office.Body;
        /**
         * Gets the date and time that an item was created. Read mode only.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
         */
        dateTimeCreated: Date;
        /**
         * Gets the date and time that an item was last modified. Read mode only.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
         *
         * Note: This member is not supported in Outlook for iOS or Outlook for Android.
         */
        dateTimeModifed: Date;
        /**
         * Gets the type of item that an instance represents.
         *
         * The itemType property returns one of the ItemType enumeration values, indicating whether the item object instance is a message or an appointment.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        itemType: Office.MailboxEnums.ItemType;
        /**
         * Gets the notification messages for an item.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        notificationMessages: Office.NotificationMessages;

       /**
        * Asynchronously loads custom properties for this add-in on the selected item.
        *
        * Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a CustomProperties object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.
        *
        * The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.
        *
        * [Api set: Mailbox 1.0]
        *
        * @remarks
        *
        * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
        *
        * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
        *
        * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
        * @param userContext - Optional. Developers can provide any object they wish to access in the callback function. This object can be accessed by the asyncResult.asyncContext property in the callback function.
        */
       loadCustomPropertiesAsync(callback: (result: AsyncResult) => void, userContext?: any): void;
    }
    /**
     * The compose mode of {@link Office.Item | Office.context.mailbox.item}.
     * 
     * Important: This is an internal Outlook object, not directly exposed through existing interfaces. You should treat this as a mode of `Office.context.mailbox.item`. Refer to the Object Model pages for more information.
     */
    export interface ItemCompose extends Item {
        /**
         * Gets or sets the description that appears in the subject field of an item.
         *
         * The subject property gets or sets the entire subject of the item, as sent by the email server.
         *
         * The subject property returns a Subject object that provides methods to get and set the subject.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        subject: Subject;
        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The addFileAttachmentAsync method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>AttachmentSizeExceeded - The attachment is larger than allowed.</td></tr><td></td><td>FileTypeNotSupported - The attachment has an extension that is not allowed.</td></tr><td></td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         * 
         * In addition to this signature, the method also has the following signatures:
         * 
         * `addFileAttachmentAsync(uri: string, attachmentName: string): void;`
         * 
         * `addFileAttachmentAsync(uri: string, attachmentName: string, options: Office.AsyncContextOptions): void;`
         * 
         * `addFileAttachmentAsync(uri: string, attachmentName: string, callback: (result: AsyncResult) => void): void;`
         *
         * @param uri - The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         *        isInline: If true, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type asyncResult. On success, the attachment identifier will be provided in the asyncResult.value property. If uploading the attachment fails, the asyncResult object will contain an Error object that provides a description of the error.
         */
        addFileAttachmentAsync(uri: string, attachmentName: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The addFileAttachmentAsync method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>AttachmentSizeExceeded - The attachment is larger than allowed.</td></tr><tr><td></td><td>FileTypeNotSupported - The attachment has an extension that is not allowed.</td></tr><tr><td></td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         *
         * @param uri - The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         */
        addFileAttachmentAsync(uri: string, attachmentName: string): void;
        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The addFileAttachmentAsync method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>AttachmentSizeExceeded - The attachment is larger than allowed.</td></tr><tr><td></td><td>FileTypeNotSupported - The attachment has an extension that is not allowed.</td></tr><tr><td></td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         *
         * @param uri - The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         *        isInline: If true, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.
         */
        addFileAttachmentAsync(uri: string, attachmentName: string, options: Office.AsyncContextOptions): void;
        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The addFileAttachmentAsync method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>AttachmentSizeExceeded - The attachment is larger than allowed.</td></tr><tr><td></td><td>FileTypeNotSupported - The attachment has an extension that is not allowed.</td></tr><tr><td></td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         *
         * @param uri - The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type asyncResult. On success, the attachment identifier will be provided in the asyncResult.value property. If uploading the attachment fails, the asyncResult object will contain an Error object that provides a description of the error.
         */
        addFileAttachmentAsync(uri: string, attachmentName: string, callback: (result: AsyncResult) => void): void;

        /**
         * Adds an Exchange item, such as a message, as an attachment to the message or appointment.
         *
         * The addItemAttachmentAsync method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, asyncResult, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the options parameter to pass state information to the callback method, if needed.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * If your Office Add-in is running in Outlook Web App, the addItemAttachmentAsync method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         * 
         * In addition to the main signature, this method also has these signatures:
         * 
         * `addItemAttachmentAsync(itemId: any, attachmentName: string): void;`
         * 
         * `addItemAttachmentAsync(itemId: any, attachmentName: string, options: Office.AsyncContextOptions): void;`
         * 
         * `addItemAttachmentAsync(itemId: any, attachmentName: string, callback: (result: AsyncResult) => void): void;`
         *
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. On success, the attachment identifier will be provided in the asyncResult.value property. If adding the attachment fails, the asyncResult object will contain an Error object that provides a description of the error.
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message or appointment.
         *
         * The addItemAttachmentAsync method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, asyncResult, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the options parameter to pass state information to the callback method, if needed.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * If your Office Add-in is running in Outlook Web App, the addItemAttachmentAsync method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         *
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string): void;
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message or appointment.
         *
         * The addItemAttachmentAsync method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, asyncResult, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the options parameter to pass state information to the callback method, if needed.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * If your Office Add-in is running in Outlook Web App, the addItemAttachmentAsync method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         *
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string, options: Office.AsyncContextOptions): void;
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message or appointment.
         *
         * The addItemAttachmentAsync method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, asyncResult, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the options parameter to pass state information to the callback method, if needed.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * If your Office Add-in is running in Outlook Web App, the addItemAttachmentAsync method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         *
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. On success, the attachment identifier will be provided in the asyncResult.value property. If adding the attachment fails, the asyncResult object will contain an Error object that provides a description of the error.
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string, callback: (result: AsyncResult) => void): void;

        /**
         * Closes the current item that is being composed
         *
         * The behaviors of the close method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or close the action.
         *
         * In the Outlook desktop client, if the message is an inline reply, the close method has no effect.
         *
         * Note: In Outlook on the web, if the item is an appointment and it has previously been saved using saveAsync, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>Restricted</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr></table>
         */
        close(): void;
        /**
         * Asynchronously returns selected data from the subject or body of a message.
         *
         * If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the InvalidSelection error.
         *
         * To access the selected data from the callback method, call asyncResult.value.data. To access the source property that the selection comes from, call asyncResult.value.sourceProperty, which will be either body or subject.
         *
         * [Api set: Mailbox 1.2]
         *
         * @returns
         * The selected data as a string with format determined by coercionType.
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr></table>
         *
         * @param coercionType - Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         */
        getSelectedDataAsync(coerciontype: Office.CoercionType, callback: (result: AsyncResult) => void): void;
        /**
         * Asynchronously returns selected data from the subject or body of a message.
         *
         * If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the InvalidSelection error.
         *
         * To access the selected data from the callback method, call asyncResult.value.data. To access the source property that the selection comes from, call asyncResult.value.sourceProperty, which will be either body or subject.
         *
         * [Api set: Mailbox 1.2]
         *
         * @returns
         * The selected data as a string with format determined by coercionType.
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr></table>
         *
         * @param coercionType - Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.
         * @param options - An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         */
        getSelectedDataAsync(coerciontype: Office.CoercionType, options: Office.AsyncContextOptions, callback: (result: AsyncResult) => void): void;
        /**
         * Removes an attachment from a message or appointment.
         *
         * The removeAttachmentAsync method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         * 
         * In addition to this signature, the method also has the following signatures:
         * 
         * `removeAttachmentAsync(attachmentIndex: string): void;`
         * 
         * `removeAttachmentAsync(attachmentIndex: string, options: Office.AsyncContextOptions): void;`
         * 
         * `removeAttachmentAsync(attachmentIndex: string, callback: (result: AsyncResult) => void): void;`
         *
         * @param attachmentIndex - The identifier of the attachment to remove. The maximum length of the string is 100 characters.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If removing the attachment fails, the asyncResult.error property will contain an error code with the reason for the failure.
         */
        removeAttachmentAsync(attachmentIndex: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Removes an attachment from a message or appointment.
         *
         * The removeAttachmentAsync method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param attachmentIndex - The identifier of the attachment to remove. The maximum length of the string is 100 characters.
         */
        removeAttachmentAsync(attachmentIndex: string): void;
        /**
         * Removes an attachment from a message or appointment.
         *
         * The removeAttachmentAsync method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param attachmentIndex - The identifier of the attachment to remove. The maximum length of the string is 100 characters.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         */
        removeAttachmentAsync(attachmentIndex: string, options: Office.AsyncContextOptions): void;
        /**
         * Removes an attachment from a message or appointment.
         *
         * The removeAttachmentAsync method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param attachmentIndex - The identifier of the attachment to remove. The maximum length of the string is 100 characters.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If removing the attachment fails, the asyncResult.error property will contain an error code with the reason for the failure.
         */
        removeAttachmentAsync(attachmentIndex: string, callback: (result: AsyncResult) => void): void;

        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.
         *
         * Since appointments have no draft state, if saveAsync is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.
         *
         * Note: If your add-in calls saveAsync on an item in compose mode in order to get an itemId to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the itemId will return an error.
         *
         * Note: The following clients have different behavior for saveAsync on appointments in compose mode:
         *
         * - Mac Outlook does not support saveAsync on a meeting in compose mode. Calling saveAsync on a meeting in Mac Outlook will return an error.
         *
         * - Outlook on the web always sends an invitation or update when saveAsync is called on an appointment in compose mode.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         * 
         * In addition to this signature, the method also has the following signatures:
         * 
         * `saveAsync(): void;`
         * 
         * `saveAsync(options: Office.AsyncContextOptions): void;`
         * 
         * `saveAsync(callback: (result: AsyncResult) => void): void;`
         *
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If removing the attachment fails, the asyncResult.error property will contain an error code with the reason for the failure.
         */
        saveAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.
         *
         * Since appointments have no draft state, if saveAsync is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.
         *
         * Note: If your add-in calls saveAsync on an item in compose mode in order to get an itemId to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the itemId will return an error.
         *
         * Note: The following clients have different behavior for saveAsync on appointments in compose mode:
         *
         * - Mac Outlook does not support saveAsync on a meeting in compose mode. Calling saveAsync on a meeting in Mac Outlook will return an error.
         *
         * - Outlook on the web always sends an invitation or update when saveAsync is called on an appointment in compose mode.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         */
        saveAsync(): void;
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.
         *
         * Since appointments have no draft state, if saveAsync is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.
         *
         * Note: If your add-in calls saveAsync on an item in compose mode in order to get an itemId to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the itemId will return an error.
         *
         * Note: The following clients have different behavior for saveAsync on appointments in compose mode:
         *
         * - Mac Outlook does not support saveAsync on a meeting in compose mode. Calling saveAsync on a meeting in Mac Outlook will return an error.
         *
         * - Outlook on the web always sends an invitation or update when saveAsync is called on an appointment in compose mode.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         */
        saveAsync(options: Office.AsyncContextOptions): void;
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.
         *
         * Since appointments have no draft state, if saveAsync is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.
         *
         * Note: If your add-in calls saveAsync on an item in compose mode in order to get an itemId to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the itemId will return an error.
         *
         * Note: The following clients have different behavior for saveAsync on appointments in compose mode:
         *
         * - Mac Outlook does not support saveAsync on a meeting in compose mode. Calling saveAsync on a meeting in Mac Outlook will return an error.
         *
         * - Outlook on the web always sends an invitation or update when saveAsync is called on an appointment in compose mode.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If removing the attachment fails, the asyncResult.error property will contain an error code with the reason for the failure.
         */
        saveAsync(callback: (result: AsyncResult) => void): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         *
         * The setSelectedDataAsync method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.
         *
         * [Api set: Mailbox 1.2]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         * 
         * In addition to this signature, the method also has the following signatures:
         * 
         * `setSelectedDataAsync(data: string): void;`
         * 
         * `setSelectedDataAsync(data: string, options: Office.AsyncContextOptions & CoercionTypeOptions): void;`
         * 
         * `setSelectedDataAsync(data: string, callback: (result: AsyncResult) => void): void;`
         *
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an ArgumentOutOfRange exception is thrown.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         *        coercionType: If text, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML. If html and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an InvalidDataFormat error is returned. If coercionType is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If removing the attachment fails, the asyncResult.error property will contain an error code with the reason for the failure.
         */
        setSelectedDataAsync(data: string, options?: Office.AsyncContextOptions & CoercionTypeOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         *
         * The setSelectedDataAsync method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.
         *
         * [Api set: Mailbox 1.2]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an ArgumentOutOfRange exception is thrown.
         */
        setSelectedDataAsync(data: string): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         *
         * The setSelectedDataAsync method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.
         *
         * [Api set: Mailbox 1.2]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an ArgumentOutOfRange exception is thrown.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         *        coercionType: If text, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML. If html and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an InvalidDataFormat error is returned. If coercionType is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.
         */
        setSelectedDataAsync(data: string, options: Office.AsyncContextOptions & CoercionTypeOptions): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         *
         * The setSelectedDataAsync method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.
         *
         * [Api set: Mailbox 1.2]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an ArgumentOutOfRange exception is thrown.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If removing the attachment fails, the asyncResult.error property will contain an error code with the reason for the failure.
         */
        setSelectedDataAsync(data: string, callback: (result: AsyncResult) => void): void;
    }
    /**
     * The read mode of {@link Office.Item | Office.context.mailbox.item}.
     * 
     * Important: This is an internal Outlook object, not directly exposed through existing interfaces. You should treat this as a mode of `Office.context.mailbox.item`. Refer to the Object Model pages for more information.
     */
    export interface ItemRead extends Item {
        /**
         * Gets an array of attachments for the item.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
         *
         * Note: Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned. For more information, see {@link https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519 | Blocked attachments in Outlook}.
         *
         */
        attachments: Office.AttachmentDetails[];
        /**
         * Gets the Exchange Web Services item class of the selected item.
         *
         *
         * You can create custom message classes that extends a default message class, for example, a custom appointment message class IPM.Appointment.Contoso.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
         * 
         * The itemClass property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.
         * 
         * <table>
         *   <tr>
         *     <th>Type</th>
         *     <th>Description</th>
         *     <th>Item Class</th>
         *   </tr>
         *   <tr>
         *     <td>Appointment items</td>
         *     <td>These are calendar items of the item class IPM.Appointment or IPM.Appointment.Occurence.</td>
         *     <td>IPM.Appointment,IPM.Appointment.Occurence</td>
         *   </tr>
         *   <tr>
         *     <td>Message items</td>
         *     <td>These include email messages that have the default message class IPM.Note, and meeting requests, responses, and cancellations, that use IPM.Schedule.Meeting as the base message class.</td>
         *     <td>IPM.Note,IPM.Schedule.Meeting.Request,IPM.Schedule.Meeting.Neg,IPM.Schedule.Meeting.Pos,IPM.Schedule.Meeting.Tent,IPM.Schedule.Meeting.Canceled</td>
         *   </tr>
         * </table>
         */
        itemClass: string;
        /**
         * Gets the Exchange Web Services item identifier for the current item.
         *
         * The itemId property is not available in compose mode. If an item identifier is required, the saveAsync method can be used to save the item to the store, which will return the item identifier in the AsyncResult.value parameter in the callback function.
         *
         * Note: The identifier returned by the itemId property is the same as the Exchange Web Services item identifier. The itemId property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API. Before making REST API calls using this value, it should be converted using Office.context.mailbox.convertToRestId. For more details, see Use the Outlook REST APIs from an Outlook add-in.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
         */
        itemId: string;
        /**
         * Gets the subject of an item, with all prefixes removed (including RE: and FWD:).
         *
         * The normalizedSubject property gets the subject of the item, with any standard prefixes (such as RE: and FW:) that are added by email programs. To get the subject of the item with the prefixes intact, use the subject property.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
         */
        normalizedSubject: string;
        /**
         * Gets the description that appears in the subject field of an item.
         *
         * The subject property gets or sets the entire subject of the item, as sent by the email server.
         *
         * The subject property returns a string. Use the normalizedSubject property to get the subject minus any leading prefixes such as RE: and FW:.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        subject: string;
        /**
         * Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.
         *
         * In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.
         *
         * If any of the string parameters exceed their limits, displayReplyAllForm throws an exception.
         *
         * When attachments are specified in the formData.attachments parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *  OR
         * An {@link Office.ReplyFormData} object that contains body or attachment data and a callback function
         */
        displayReplyAllForm(formData: string | ReplyFormData): void;
        /**
         * Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.
         *
         * In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.
         *
         * If any of the string parameters exceed their limits, displayReplyForm throws an exception.
         *
         * When attachments are specified in the formData.attachments parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.
         * OR
         * An {@link Office.ReplyFormData} object that contains body or attachment data and a callback function.
         */
        displayReplyForm(formData: string | ReplyFormData): void;
        /**
         * Gets the entities found in the selected item.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
         */
        getEntities(): Entities;
        /**
         * Gets an array of all the entities of the specified entity type found in the selected item.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         * 
         * @param entityType - One of the EntityType enumeration values.
         *
         * @returns
         * If the value passed in entityType is not a valid member of the EntityType enumeration, the method returns null. If no entities of the specified type are present on the item, the method returns an empty array. Otherwise, the type of the objects in the returned array depends on the type of entity requested in the entityType parameter.
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>Restricted</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
         * 
         * While the minimum permission level to use this method is Restricted, some entity types require ReadItem to access, as specified in the following table.
         * 
         * <table>
         *   <tr>
         *     <th>Value of entityType</th>
         *     <th>Type of objects in returned array</th>
         *     <th>Required Permission Leve</th>
         *   </tr>
         *   <tr>
         *     <td>Address</td>
         *     <td>String</td>
         *     <td>Restricted</td>
         *   </tr>
         *   <tr>
         *     <td>Contact</td>
         *     <td>Contact</td>
         *     <td>ReadItem</td>
         *   </tr>
         *   <tr>
         *     <td>EmailAddress</td>
         *     <td>String</td>
         *     <td>ReadItem</td>
         *   </tr>
         *   <tr>
         *     <td>MeetingSuggestion</td>
         *     <td>MeetingSuggestion</td>
         *     <td>ReadItem</td>
         *   </tr>
         *   <tr>
         *     <td>PhoneNumber</td>
         *     <td>PhoneNumber</td>
         *     <td>Restricted</td>
         *   </tr>
         *   <tr>
         *     <td>TaskSuggestion</td>
         *     <td>TaskSuggestion</td>
         *     <td>ReadItem</td>
         *   </tr>
         *   <tr>
         *     <td>URL</td>
         *     <td>String</td>
         *     <td>Restricted</td>
         *   </tr>
         * </table>
         */
        getEntitiesByType(entityType: Office.MailboxEnums.EntityType): (string | Contact | MeetingSuggestion | PhoneNumber | TaskSuggestion)[];
        /**
         * Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.
         *
         * The getFilteredEntitiesByName method returns the entities that match the regular expression defined in the ItemHasKnownEntity rule element in the manifest XML file with the specified FilterName element value.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
         *
         * @param name - The name of the ItemHasKnownEntity rule element that defines the filter to match.
         * @returns If there is no ItemHasKnownEntity element in the manifest with a FilterName element value that matches the name parameter, the method returns null. If the name parameter does match an ItemHasKnownEntity element in the manifest, but there are no entities in the current item that match, the method return an empty array.
         */
        getFilteredEntitiesByName(name: string): (string | Contact | MeetingSuggestion | PhoneNumber | TaskSuggestion)[];
        /**
         * Returns string values in the selected item that match the regular expressions defined in the manifest XML file.
         *
         * The getRegExMatches method returns the strings that match the regular expression defined in each ItemHasRegularExpressionMatch or ItemHasKnownEntity rule element in the manifest XML file. For an ItemHasRegularExpressionMatch rule, a matching string has to occur in the property of the item that is specified by that rule. The PropertyName simple type defines the supported properties.
         *
         * If you specify an ItemHasRegularExpressionMatch rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as .* to obtain the entire body of an item does not always return the expected results. Instead, use the Body.getAsync method to retrieve the entire body.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @returns
         * An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the RegExName attribute of the matching ItemHasRegularExpressionMatch rule or the FilterName attribute of the matching ItemHasKnownEntity rule.
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
         */
        getRegExMatches(): any;
        /**
         * Returns string values in the selected item that match the named regular expression defined in the manifest XML file.
         *
         * The getRegExMatchesByName method returns the strings that match the regular expression defined in the ItemHasRegularExpressionMatch rule element in the manifest XML file with the specified RegExName element value.
         *
         * If you specify an ItemHasRegularExpressionMatch rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as .* to obtain the entire body of an item does not always return the expected results.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @returns
         * An array that contains the strings that match the regular expression defined in the manifest XML file.
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
         *
         * @param name - The name of the ItemHasRegularExpressionMatch rule element that defines the filter to match.
         */
        getRegExMatchesByName(name: string): string[];
    }
    /**
     * A subclass of {@link Office.Item} for messages.
     * 
     * Important: This is an internal Outlook object, not directly exposed through existing interfaces. You should treat this as a mode of `Office.context.mailbox.item`. Refer to the Object Model pages for more information.
     */
    export interface Message extends Item {
        /**
         * Gets an identifier for the email conversation that contains a particular message.
         *
         * You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.
         *
         * You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the conversationId property will return a value.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        conversationId: string;
    }

     /**
     * The message compose mode of {@link Office.Item | Office.context.mailbox.item}.
     * 
     * Important: This is an internal Outlook object, not directly exposed through existing interfaces. You should treat this as a mode of `Office.context.mailbox.item`. Refer to the Object Model pages for more information.
     */
    export interface MessageCompose extends Message, ItemCompose {
        /**
         * Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr></table>
         */
        bcc: Recipients;
        /**
         * Gets an object that provides methods for manipulating the body of an item.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr></table>
         */
        body: Office.Body;
        /**
         * Provides access to the Cc (carbon copy) recipients of a message. The type of object and level of access depends on the mode of the current item.
         *
         * The cc property returns a {@link Office.Recipients} object that provides methods to get or update the recipients on the Cc line of the message.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr></table>
         */
        cc: Recipients;
        /**
         * Gets an identifier for the email conversation that contains a particular message.
         *
         * You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.
         *
         * You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the conversationId property will return a value.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr></table>
         */
        conversationId: string;
        /**
         * Gets the date and time that an item was created. Read mode only.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr></table>
         */
        dateTimeCreated: Date;
        /**
         * Gets the date and time that an item was last modified. Read mode only.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr></table>
         *
         * Note: This member is not supported in Outlook for iOS or Outlook for Android.
         */
        dateTimeModifed: Date;
        /**
         * Gets the type of item that an instance represents.
         *
         * The itemType property returns one of the ItemType enumeration values, indicating whether the item object instance is a message or an appointment.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr></table>
         */
        itemType: Office.MailboxEnums.ItemType;
        /**
         * Gets the notification messages for an item.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr></table>
         */
        notificationMessages: Office.NotificationMessages;
        /**
         * Gets or sets the description that appears in the subject field of an item.
         *
         * The subject property gets or sets the entire subject of the item, as sent by the email server.
         *
         * The subject property returns a Subject object that provides methods to get and set the subject.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr></table>
         */
        subject: Subject;
        /**
         * Provides access to the recipients on the To line of a message. The type of object and level of access depends on the mode of the current item.
         *
         * The to property returns a Recipients object that provides methods to get or update the recipients on the To line of the message.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr></table>
         */
        to: Recipients;

        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The addFileAttachmentAsync method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr>
         *
         * <tr><td>Errors</td><td>AttachmentSizeExceeded - The attachment is larger than allowed.</td></tr><tr><td></td><td>FileTypeNotSupported - The attachment has an extension that is not allowed.</td></tr><tr><td></td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         * 
         * In addition to this signature, the method also has the following signatures:
         * 
         * `addFileAttachmentAsync(uri: string, attachmentName: string): void;`
         * 
         * `addFileAttachmentAsync(uri: string, attachmentName: string, options: AsyncContextOptions): void;`
         * 
         * `addFileAttachmentAsync(uri: string, attachmentName: string, callback: (result: AsyncResult) => void): void;`
         *
         * @param uri - The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         *        isInline: If true, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type asyncResult. On success, the attachment identifier will be provided in the asyncResult.value property. If uploading the attachment fails, the asyncResult object will contain an Error object that provides a description of the error.
         */
        addFileAttachmentAsync(uri: string, attachmentName: string, options?: AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The addFileAttachmentAsync method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr>
         *
         * <tr><td>Errors</td><td>AttachmentSizeExceeded - The attachment is larger than allowed.</td></tr><tr><td></td><td>FileTypeNotSupported - The attachment has an extension that is not allowed.</td></tr><tr><td></td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         *
         * @param uri - The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         */
        addFileAttachmentAsync(uri: string, attachmentName: string): void;
        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The addFileAttachmentAsync method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr>
         *
         * <tr><td>Errors</td><td>AttachmentSizeExceeded - The attachment is larger than allowed.</td></tr><tr><td></td><td>FileTypeNotSupported - The attachment has an extension that is not allowed.</td></tr><tr><td></td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         *
         * @param uri - The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         *        isInline: If true, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.
         */
        addFileAttachmentAsync(uri: string, attachmentName: string, options: AsyncContextOptions): void;
        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The addFileAttachmentAsync method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr>
         *
         * <tr><td>Errors</td><td>AttachmentSizeExceeded - The attachment is larger than allowed.</td></tr><tr><td></td><td>FileTypeNotSupported - The attachment has an extension that is not allowed.</td></tr><tr><td></td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         *
         * @param uri - The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type asyncResult. On success, the attachment identifier will be provided in the asyncResult.value property. If uploading the attachment fails, the asyncResult object will contain an Error object that provides a description of the error.
         */
        addFileAttachmentAsync(uri: string, attachmentName: string, callback: (result: AsyncResult) => void): void;
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message or appointment.
         *
         * The addItemAttachmentAsync method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, asyncResult, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the options parameter to pass state information to the callback method, if needed.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * If your Office Add-in is running in Outlook Web App, the addItemAttachmentAsync method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr>
         *
         * <tr><td>Errors</td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         * 
         * In addition to the main signature, this method also has these signatures:
         * 
         * `addItemAttachmentAsync(itemId: any, attachmentName: string): void;`
         * 
         * `addItemAttachmentAsync(itemId: any, attachmentName: string, options: Office.AsyncContextOptions): void;`
         * 
         * `addItemAttachmentAsync(itemId: any, attachmentName: string, callback: (result: AsyncResult) => void): void;`
         *
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. On success, the attachment identifier will be provided in the asyncResult.value property. If adding the attachment fails, the asyncResult object will contain an Error object that provides a description of the error.
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message or appointment.
         *
         * The addItemAttachmentAsync method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, asyncResult, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the options parameter to pass state information to the callback method, if needed.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * If your Office Add-in is running in Outlook Web App, the addItemAttachmentAsync method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr>
         *
         * <tr><td>Errors</td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         *
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string): void;
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message or appointment.
         *
         * The addItemAttachmentAsync method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, asyncResult, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the options parameter to pass state information to the callback method, if needed.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * If your Office Add-in is running in Outlook Web App, the addItemAttachmentAsync method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr>
         *
         * <tr><td>Errors</td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         *
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string, options: Office.AsyncContextOptions): void;
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message or appointment.
         *
         * The addItemAttachmentAsync method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, asyncResult, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the options parameter to pass state information to the callback method, if needed.
         *
         * You can subsequently use the identifier with the removeAttachmentAsync method to remove the attachment in the same session.
         *
         * If your Office Add-in is running in Outlook Web App, the addItemAttachmentAsync method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr>
         *
         * <tr><td>Errors</td><td>NumberOfAttachmentsExceeded - The message or appointment has too many attachments.</td></tr></table>
         *
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. On success, the attachment identifier will be provided in the asyncResult.value property. If adding the attachment fails, the asyncResult object will contain an Error object that provides a description of the error.
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string, callback: (result: AsyncResult) => void): void;
        /**
         * Closes the current item that is being composed
         *
         * The behaviors of the close method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or close the action.
         *
         * In the Outlook desktop client, if the message is an inline reply, the close method has no effect.
         *
         * Note: In Outlook on the web, if the item is an appointment and it has previously been saved using saveAsync, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>Restricted</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr></table>
         */
        close(): void;
        /**
         * Asynchronously returns selected data from the subject or body of a message.
         *
         * If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the InvalidSelection error.
         *
         * To access the selected data from the callback method, call asyncResult.value.data. To access the source property that the selection comes from, call asyncResult.value.sourceProperty, which will be either body or subject.
         *
         * [Api set: Mailbox 1.2]
         *
         * @returns
         * The selected data as a string with format determined by coercionType.
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr></table>
         *
         * @param coercionType - Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         */
        getSelectedDataAsync(coerciontype: Office.CoercionType, callback: (result: AsyncResult) => void): void;
        /**
         * Asynchronously returns selected data from the subject or body of a message.
         *
         * If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the InvalidSelection error.
         *
         * To access the selected data from the callback method, call asyncResult.value.data. To access the source property that the selection comes from, call asyncResult.value.sourceProperty, which will be either body or subject.
         *
         * [Api set: Mailbox 1.2]
         *
         * @returns
         * The selected data as a string with format determined by coercionType.
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr></table>
         *
         * @param coercionType - Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.
         * @param options - An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         */
        getSelectedDataAsync(coerciontype: Office.CoercionType, options: Office.AsyncContextOptions, callback: (result: AsyncResult) => void): void;
        /**
         * Asynchronously loads custom properties for this add-in on the selected item.
         *
         * Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a CustomProperties object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.
         *
         * The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr></table>
         *
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         * @param userContext - Optional. Developers can provide any object they wish to access in the callback function. This object can be accessed by the asyncResult.asyncContext property in the callback function.
         */
        loadCustomPropertiesAsync(callback: (result: AsyncResult) => void, userContext?: any): void;
        /**
         * Removes an attachment from a message or appointment.
         *
         * The removeAttachmentAsync method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         * 
         * In addition to this signature, the method also has the following signatures:
         * 
         * `removeAttachmentAsync(attachmentIndex: string): void;`
         * 
         * `removeAttachmentAsync(attachmentIndex: string, options: Office.AsyncContextOptions): void;`
         * 
         * `removeAttachmentAsync(attachmentIndex: string, callback: (result: AsyncResult) => void): void;`
         *
         * @param attachmentIndex - The identifier of the attachment to remove. The maximum length of the string is 100 characters.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If removing the attachment fails, the asyncResult.error property will contain an error code with the reason for the failure.
         */
        removeAttachmentAsync(attachmentIndex: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Removes an attachment from a message or appointment.
         *
         * The removeAttachmentAsync method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param attachmentIndex - The identifier of the attachment to remove. The maximum length of the string is 100 characters.
         */
        removeAttachmentAsync(attachmentIndex: string): void;
        /**
         * Removes an attachment from a message or appointment.
         *
         * The removeAttachmentAsync method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param attachmentIndex - The identifier of the attachment to remove. The maximum length of the string is 100 characters.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         */
        removeAttachmentAsync(attachmentIndex: string, options: Office.AsyncContextOptions): void;
        /**
         * Removes an attachment from a message or appointment.
         *
         * The removeAttachmentAsync method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param attachmentIndex - The identifier of the attachment to remove. The maximum length of the string is 100 characters.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If removing the attachment fails, the asyncResult.error property will contain an error code with the reason for the failure.
         */
        removeAttachmentAsync(attachmentIndex: string, callback: (result: AsyncResult) => void): void;
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.
         *
         * Since appointments have no draft state, if saveAsync is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.
         *
         * Note: If your add-in calls saveAsync on an item in compose mode in order to get an itemId to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the itemId will return an error.
         *
         * Note: The following clients have different behavior for saveAsync on appointments in compose mode:
         *
         * - Mac Outlook does not support saveAsync on a meeting in compose mode. Calling saveAsync on a meeting in Mac Outlook will return an error.
         *
         * - Outlook on the web always sends an invitation or update when saveAsync is called on an appointment in compose mode.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         * 
         * In addition to this signature, the method also has the following signatures:
         * 
         * `saveAsync(): void;`
         * 
         * `saveAsync(options: Office.AsyncContextOptions): void;`
         * 
         * `saveAsync(callback: (result: AsyncResult) => void): void;`
         *
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If removing the attachment fails, the asyncResult.error property will contain an error code with the reason for the failure.
         */
        saveAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.
         *
         * Since appointments have no draft state, if saveAsync is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.
         *
         * Note: If your add-in calls saveAsync on an item in compose mode in order to get an itemId to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the itemId will return an error.
         *
         * Note: The following clients have different behavior for saveAsync on appointments in compose mode:
         *
         * - Mac Outlook does not support saveAsync on a meeting in compose mode. Calling saveAsync on a meeting in Mac Outlook will return an error.
         *
         * - Outlook on the web always sends an invitation or update when saveAsync is called on an appointment in compose mode.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         */
        saveAsync(): void;
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.
         *
         * Since appointments have no draft state, if saveAsync is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.
         *
         * Note: If your add-in calls saveAsync on an item in compose mode in order to get an itemId to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the itemId will return an error.
         *
         * Note: The following clients have different behavior for saveAsync on appointments in compose mode:
         *
         * - Mac Outlook does not support saveAsync on a meeting in compose mode. Calling saveAsync on a meeting in Mac Outlook will return an error.
         *
         * - Outlook on the web always sends an invitation or update when saveAsync is called on an appointment in compose mode.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         */
        saveAsync(options: Office.AsyncContextOptions): void;
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.
         *
         * Since appointments have no draft state, if saveAsync is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.
         *
         * Note: If your add-in calls saveAsync on an item in compose mode in order to get an itemId to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the itemId will return an error.
         *
         * Note: The following clients have different behavior for saveAsync on appointments in compose mode:
         *
         * - Mac Outlook does not support saveAsync on a meeting in compose mode. Calling saveAsync on a meeting in Mac Outlook will return an error.
         *
         * - Outlook on the web always sends an invitation or update when saveAsync is called on an appointment in compose mode.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If removing the attachment fails, the asyncResult.error property will contain an error code with the reason for the failure.
         */
        saveAsync(callback: (result: AsyncResult) => void): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         *
         * The setSelectedDataAsync method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.
         *
         * [Api set: Mailbox 1.2]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         * 
         * In addition to this signature, the method also has the following signatures:
         * 
         * `setSelectedDataAsync(data: string): void;`
         * 
         * `setSelectedDataAsync(data: string, options: Office.AsyncContextOptions & CoercionTypeOptions): void;`
         * 
         * `setSelectedDataAsync(data: string, callback: (result: AsyncResult) => void): void;`
         *
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an ArgumentOutOfRange exception is thrown.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         *        coercionType: If text, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML. If html and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an InvalidDataFormat error is returned. If coercionType is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If removing the attachment fails, the asyncResult.error property will contain an error code with the reason for the failure.
         */
        setSelectedDataAsync(data: string, options?: Office.AsyncContextOptions & CoercionTypeOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         *
         * The setSelectedDataAsync method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.
         *
         * [Api set: Mailbox 1.2]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an ArgumentOutOfRange exception is thrown.
         */
        setSelectedDataAsync(data: string): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         *
         * The setSelectedDataAsync method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.
         *
         * [Api set: Mailbox 1.2]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an ArgumentOutOfRange exception is thrown.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         *        coercionType: If text, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML. If html and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an InvalidDataFormat error is returned. If coercionType is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.
         */
        setSelectedDataAsync(data: string, options: Office.AsyncContextOptions & CoercionTypeOptions): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         *
         * The setSelectedDataAsync method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.
         *
         * [Api set: Mailbox 1.2]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidAttachmentId - The attachment identifier does not exist.</td></tr></table>
         *
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an ArgumentOutOfRange exception is thrown.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If removing the attachment fails, the asyncResult.error property will contain an error code with the reason for the failure.
         */
        setSelectedDataAsync(data: string, callback: (result: AsyncResult) => void): void;
    }
    /**
     * The message read mode of {@link Office.Item | Office.context.mailbox.item}.
     * 
     * Important: This is an internal Outlook object, not directly exposed through existing interfaces. You should treat this as a mode of Office.context.mailbox.item. Refer to the Object Model pages for more information.
     */
    export interface MessageRead extends Message, ItemRead {
        /**
         * Gets an array of attachments for the item.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
         * 
         * Note: Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned. For more information, see {@link https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519 | Blocked attachments in Outlook}.
         *
         */
        attachments: Office.AttachmentDetails[];
        /**
         * Gets an object that provides methods for manipulating the body of an item.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
         */
        body: Office.Body;
        /**
         * Provides access to the Cc (carbon copy) recipients of a message. The type of object and level of access depends on the mode of the current item.
         *
         * The cc property returns an array that contains an EmailAddressDetails object for each recipient listed on the Cc line of the message. The collection is limited to a maximum of 100 members.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
         */
        cc: EmailAddressDetails[];
        /**
         * Gets an identifier for the email conversation that contains a particular message.
         *
         * You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.
         *
         * You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the conversationId property will return a value.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
         */
        conversationId: string;
        /**
         * Gets the date and time that an item was created. Read mode only.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
         */
        dateTimeCreated: Date;
        /**
         * Gets the date and time that an item was last modified. Read mode only.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
         *
         * Note: This member is not supported in Outlook for iOS or Outlook for Android.
         */
        dateTimeModifed: Date;
        /**
         * Gets the email address of the sender of a message.
         *
         * The from and sender properties represent the same person unless the message is sent by a delegate. In that case, the from property represents the delegator, and the sender property represents the delegate.
         *
         * Note: The recipientType property of the EmailAddressDetails object in the from property is undefined.
         * 
         * The from property returns an EmailAddressDetails object.
         * 
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
         */
        from: EmailAddressDetails;
        /**
         * Gets the Internet message identifier for an email message.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
         */
        internetMessageId: string;
        /**
         * Gets the Exchange Web Services item class of the selected item.
         * 
         * You can create custom message classes that extends a default message class, for example, a custom appointment message class IPM.Appointment.Contoso.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * 
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
		 
         * The itemClass property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.
         * 
         * <table>
         *   <tr>
         *     <th>Type</th>
         *     <th>Description</th>
         *     <th>Item Class</th>
         *   </tr>
         *   <tr>
         *     <td>Appointment items</td>
         *     <td>These are calendar items of the item class IPM.Appointment or IPM.Appointment.Occurence.</td>
         *     <td>IPM.Appointment,IPM.Appointment.Occurence</td>
         *   </tr>
         *   <tr>
         *     <td>Message items</td>
         *     <td>These include email messages that have the default message class IPM.Note, and meeting requests, responses, and cancellations, that use IPM.Schedule.Meeting as the base message class.</td>
         *     <td>IPM.Note,IPM.Schedule.Meeting.Request,IPM.Schedule.Meeting.Neg,IPM.Schedule.Meeting.Pos,IPM.Schedule.Meeting.Tent,IPM.Schedule.Meeting.Canceled</td>
         *   </tr>
         * </table>
         * 
         */
        itemClass: string;
        /**
         * Gets the Exchange Web Services item identifier for the current item.
         *
         * The itemId property is not available in compose mode. If an item identifier is required, the saveAsync method can be used to save the item to the store, which will return the item identifier in the AsyncResult.value parameter in the callback function.
         *
         * Note: The identifier returned by the itemId property is the same as the Exchange Web Services item identifier. The itemId property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API. Before making REST API calls using this value, it should be converted using Office.context.mailbox.convertToRestId. For more details, see Use the Outlook REST APIs from an Outlook add-in.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
         */
        itemId: string;
        /**
         * Gets the type of item that an instance represents.
         *
         * The itemType property returns one of the ItemType enumeration values, indicating whether the item object instance is a message or an appointment.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
         */
        itemType: Office.MailboxEnums.ItemType;
        /**
         * Gets the subject of an item, with all prefixes removed (including RE: and FWD:).
         *
         * The normalizedSubject property gets the subject of the item, with any standard prefixes (such as RE: and FW:) that are added by email programs. To get the subject of the item with the prefixes intact, use the subject property.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
         */
        normalizedSubject: string;
        /**
         * Gets the notification messages for an item.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        notificationMessages: Office.NotificationMessages;
        /**
         * Gets the email address of the sender of an email message.
         *
         * The from and sender properties represent the same person unless the message is sent by a delegate. In that case, the from property represents the delegator, and the sender property represents the delegate.
         *
         * Note: The recipientType property of the EmailAddressDetails object in the sender property is undefined.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
         */
        sender: EmailAddressDetails;
        /**
         * Gets the description that appears in the subject field of an item.
         *
         * The subject property gets or sets the entire subject of the item, as sent by the email server.
         *
         * The subject property returns a string. Use the normalizedSubject property to get the subject minus any leading prefixes such as RE: and FW:.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
         */
        subject: string;
        /**
         * Provides access to the recipients on the To line of a message. The type of object and level of access depends on the mode of the current item.
         *
         * The to property returns an array that contains an EmailAddressDetails object for each recipient listed on the To line of the message. The collection is limited to a maximum of 100 members.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
         */
        to: EmailAddressDetails[];
        /**
         * Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.
         *
         * In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.
         *
         * If any of the string parameters exceed their limits, displayReplyAllForm throws an exception.
         *
         * When attachments are specified in the formData.attachments parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *  OR
         * An {@link Office.ReplyFormData} object that contains body or attachment data and a callback function
         */
        displayReplyAllForm(formData: string | ReplyFormData): void;
        /**
         * Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.
         *
         * In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.
         *
         * If any of the string parameters exceed their limits, displayReplyForm throws an exception.
         *
         * When attachments are specified in the formData.attachments parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.
         * OR
         * An {@link Office.ReplyFormData} object that contains body or attachment data and a callback function.
         */
        displayReplyForm(formData: string | ReplyFormData): void;
        /**
         * Gets the entities found in the selected item.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
         */
        getEntities(): Entities;
        /**
         * Gets an array of all the entities of the specified entity type found in the selected item.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @param entityType - One of the EntityType enumeration values.
         * 
         * @returns
         * If the value passed in entityType is not a valid member of the EntityType enumeration, the method returns null. If no entities of the specified type are present on the item, the method returns an empty array. Otherwise, the type of the objects in the returned array depends on the type of entity requested in the entityType parameter.
         *
         * @remarks
         * 
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>Restricted</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
         * 
         * While the minimum permission level to use this method is Restricted, some entity types require ReadItem to access, as specified in the following table.
         * 
         * <table>
         *   <tr>
         *     <th>Value of entityType</th>
         *     <th>Type of objects in returned array</th>
         *     <th>Required Permission Leve</th>
         *   </tr>
         *   <tr>
         *     <td>Address</td>
         *     <td>String</td>
         *     <td>Restricted</td>
         *   </tr>
         *   <tr>
         *     <td>Contact</td>
         *     <td>Contact</td>
         *     <td>ReadItem</td>
         *   </tr>
         *   <tr>
         *     <td>EmailAddress</td>
         *     <td>String</td>
         *     <td>ReadItem</td>
         *   </tr>
         *   <tr>
         *     <td>MeetingSuggestion</td>
         *     <td>MeetingSuggestion</td>
         *     <td>ReadItem</td>
         *   </tr>
         *   <tr>
         *     <td>PhoneNumber</td>
         *     <td>PhoneNumber</td>
         *     <td>Restricted</td>
         *   </tr>
         *   <tr>
         *     <td>TaskSuggestion</td>
         *     <td>TaskSuggestion</td>
         *     <td>ReadItem</td>
         *   </tr>
         *   <tr>
         *     <td>URL</td>
         *     <td>String</td>
         *     <td>Restricted</td>
         *   </tr>
         * </table>
         */
        getEntitiesByType(entityType: Office.MailboxEnums.EntityType): (string | Contact | MeetingSuggestion | PhoneNumber | TaskSuggestion)[];
        /**
         * Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.
         *
         * The getFilteredEntitiesByName method returns the entities that match the regular expression defined in the ItemHasKnownEntity rule element in the manifest XML file with the specified FilterName element value.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
         *
         * @param name - The name of the ItemHasKnownEntity rule element that defines the filter to match.
         * @returns If there is no ItemHasKnownEntity element in the manifest with a FilterName element value that matches the name parameter, the method returns null. If the name parameter does match an ItemHasKnownEntity element in the manifest, but there are no entities in the current item that match, the method return an empty array.
         */
        getFilteredEntitiesByName(name: string): (string | Contact | MeetingSuggestion | PhoneNumber | TaskSuggestion)[];
        /**
         * Returns string values in the selected item that match the regular expressions defined in the manifest XML file.
         *
         * The getRegExMatches method returns the strings that match the regular expression defined in each ItemHasRegularExpressionMatch or ItemHasKnownEntity rule element in the manifest XML file. For an ItemHasRegularExpressionMatch rule, a matching string has to occur in the property of the item that is specified by that rule. The PropertyName simple type defines the supported properties.
         *
         * If you specify an ItemHasRegularExpressionMatch rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as .* to obtain the entire body of an item does not always return the expected results. Instead, use the Body.getAsync method to retrieve the entire body.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @returns
         * An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the RegExName attribute of the matching ItemHasRegularExpressionMatch rule or the FilterName attribute of the matching ItemHasKnownEntity rule.
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
         */
        getRegExMatches(): any;
        /**
         * Returns string values in the selected item that match the named regular expression defined in the manifest XML file.
         *
         * The getRegExMatchesByName method returns the strings that match the regular expression defined in the ItemHasRegularExpressionMatch rule element in the manifest XML file with the specified RegExName element value.
         *
         * If you specify an ItemHasRegularExpressionMatch rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as .* to obtain the entire body of an item does not always return the expected results.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @returns
         * An array that contains the strings that match the regular expression defined in the manifest XML file.
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
         *
         * @param name - The name of the ItemHasRegularExpressionMatch rule element that defines the filter to match.
         */
        getRegExMatchesByName(name: string): string[];
        /**
         * Asynchronously loads custom properties for this add-in on the selected item.
         *
         * Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a CustomProperties object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.
         *
         * The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Message Read</td></tr></table>
         *
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         * @param userContext - Optional. Developers can provide any object they wish to access in the callback function. This object can be accessed by the asyncResult.asyncContext property in the callback function.
         */
        loadCustomPropertiesAsync(callback: (result: AsyncResult) => void, userContext?: any): void;
    }

    /**
     * Represents a date and time in the local client's time zone. Read mode only.
     *
     * [Api set: Mailbox 1.0]
     *
     * @remarks
     *
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
     */
    export interface LocalClientTime {
        /**
         * Integer value representing the month, beginning with 0 for January to 11 for December.
         */
        month: number;
        /**
         * Integer value representing the day of the month.
         */
        date: number;
        /**
         * Integer value repesenting the year.
         */
        year: number;
        /**
         * Integer value representing the hour on a 24-hour clock.
         */
        hours: number;
        /**
         * Integer value representing the minutes.
         */
        minutes: number;
        /**
         * Integer value representing the seconds.
         */
        seconds: number;
        /**
         * Integer value representing the milliseconds.
         */
        milliseconds: number;
        /**
         * Integer value representing the number of minutes difference between the local time zone and UTC.
         */
        timezoneOffset: number;
    }
    /**
     * Provides methods to get and set the location of a meeting in an Outlook add-in.
     *
     * [Api set: Mailbox 1.1]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr></table>
     */
    export interface Location {
        /**
         * Gets the location of an appointment.
         *
         * The getAsync method starts an asynchronous call to the Exchange server to get the location of an appointment. The location of the appointment is provided as a string in the asyncResult.value property.
         *
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr></table>
         * 
         * In addition to this signature, the method also has the following signature:
         * 
         * `getAsync(callback: (result: AsyncResult) => void): void;`
         * 
         */
        getAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Gets the location of an appointment.
         *
         * The getAsync method starts an asynchronous call to the Exchange server to get the location of an appointment. The location of the appointment is provided as a string in the asyncResult.value property.
         *
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr></table>
         */
        getAsync(callback: (result: AsyncResult) => void): void;
        /**
         * Sets the location of an appointment.
         *
         * The setAsync method starts an asynchronous call to the Exchange server to set the location of an appointment. Setting the location of an appointment overwrites the current location.
         *
         * @param location - The location of the appointment. The string is limited to 255 characters.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.  If setting the location fails, the asyncResult.error property will contain an error code.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>DataExceedsMaximumSize - The location parameter is longer than 255 characters.</td></tr></table>
         * 
         * In addition to this signature, the method also has the following signatures:
         * 
         * `setAsync(location: string): void;`
         * 
         * `setAsync(location: string, options: Office.AsyncContextOptions): void;`
         * 
         * `setAsync(location: string, callback: (result: AsyncResult) => void): void;`
         */
        setAsync(location: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Sets the location of an appointment.
         *
         * The setAsync method starts an asynchronous call to the Exchange server to set the location of an appointment. Setting the location of an appointment overwrites the current location.
         *
         * @param location - The location of the appointment. The string is limited to 255 characters.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>DataExceedsMaximumSize - The location parameter is longer than 255 characters.</td></tr></table>
         */
        setAsync(location: string): void;
        /**
         * Sets the location of an appointment.
         *
         * The setAsync method starts an asynchronous call to the Exchange server to set the location of an appointment. Setting the location of an appointment overwrites the current location.
         *
         * @param location - The location of the appointment. The string is limited to 255 characters.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>DataExceedsMaximumSize - The location parameter is longer than 255 characters.</td></tr></table>
         */
        setAsync(location: string, options: Office.AsyncContextOptions): void;
        /**
         * Sets the location of an appointment.
         *
         * The setAsync method starts an asynchronous call to the Exchange server to set the location of an appointment. Setting the location of an appointment overwrites the current location.
         *
         * @param location - The location of the appointment. The string is limited to 255 characters.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.  If setting the location fails, the asyncResult.error property will contain an error code.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>DataExceedsMaximumSize - The location parameter is longer than 255 characters.</td></tr></table>
         */
        setAsync(location: string, callback: (result: AsyncResult) => void): void;
    }
    /**
     * Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.
     *
     * Namespaces:
     *
     * - diagnostics: Provides diagnostic information to an Outlook add-in.
     *
     * - item: Provides methods and properties for accessing a message or appointment in an Outlook add-in.
     *
     * - userProfile: Provides information about the user in an Outlook add-in.
     *
     * [Api set: Mailbox 1.0]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>Restricted</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
     */
    export interface Mailbox {
        /**
         * Provides diagnostic information to an Outlook add-in.
         * 
         * Contains the following members:
         * 
         *  - hostName (string): A string that represents the name of the host application. It be one of the following values: Outlook, Mac Outlook, OutlookIOS, or OutlookWebApp.
         * 
         *  - hostVersion (string): A string that represents the version of either the host application or the Exchange Server. If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the hostVersion property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string 15.0.468.0.
         * 
         *  - OWAView (string): A string that represents the current view of Outlook Web App. If the host application is not Outlook Web App, then accessing this property results in undefined. Outlook Web App has three views (OneColumn - displayed when the screen is narrow, TwoColumns - displayed when the screen is wider, and ThreeColumns - displayed when the screen is wide.) that correspond to the width of the screen and the window, and the number of columns that can be displayed.
         *
         *  More information is under {@link Office.Diagnostics}. 
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        diagnostics: Office.Diagnostics;
        /**
         * Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.
         *
         * Your app must have the ReadItem permission specified in its manifest to call the ewsUrl member in read mode.
         *
         * In compose mode you must call the saveAsync method before you can use the ewsUrl member. Your app must have ReadWriteItem permissions to call the saveAsync method.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         *
         * The ewsUrl value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to {@link https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item | get attachments from the selected item}.
         *
         * Note: This member is not supported in Outlook for iOS or Outlook for Android.
         */
        ewsUrl: string;
        /**
         * The mailbox item.  Depending on the context in which the add-in opened, the item may be of any number of types.
         * If you want to see IntelliSense for only a specific type, you should cast this item to one of the following:
         * `ItemCompose`, `ItemRead`, `MessageCompose`, `MessageRead`, `AppointmentCompose`, `AppointmentRead`
         */
        item: Item & ItemCompose & ItemRead & MessageRead & MessageCompose & AppointmentRead & AppointmentCompose;
        /**
         * Information about the user associated with the mailbox. This includes their account type, display name, email adddress, and time zone.
         * 
         * More information is under {@link Office.UserProfile}
         */
        userProfile: Office.UserProfile;
        /**
         * Converts an item ID formatted for REST into EWS format.
         *
         * Item IDs retrieved via a REST API (such as the Outlook Mail API or the Microsoft Graph) use a different format than the format used by Exchange Web Services (EWS). The convertToEwsId method converts a REST-formatted ID into the proper format for EWS.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>Restricted</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         *
         * @param itemId - An item ID formatted for the Outlook REST APIs.
         * @param restVersion - A value indicating the version of the Outlook REST API used to retrieve the item ID.
         */
        convertToEwsId(itemId: string, restVersion: Office.MailboxEnums.RestVersion): string;
        /**
         * Gets a dictionary containing time information in local client time.
         *
         * The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.
         *
         * If the mail app is running in Outlook, the convertToLocalClientTime method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the convertToLocalClientTime method will return a dictionary object with the values set to the time zone specified in the EAC.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         *
         * @param timeValue - A Date object.
         */
        convertToLocalClientTime(timeValue: Date): LocalClientTime;
        /**
         * Converts an item ID formatted for EWS into REST format.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * 
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>Restricted</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         *
         * Item IDs retrieved via EWS or via the itemId property use a different format than the format used by REST APIs (such as the {@link https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations | Outlook Mail API} or the {@link http://graph.microsoft.io/ | Microsoft Graph}. The convertToRestId method converts an EWS-formatted ID into the proper format for REST.
         *
         * @param itemId - An item ID formatted for Exchange Web Services (EWS)
         * @param restVersion - A value indicating the version of the Outlook REST API that the converted ID will be used with.
         */
        convertToRestId(itemId: string, restVersion: Office.MailboxEnums.RestVersion): string;
        /**
         * Gets a Date object from a dictionary containing time information.
         *
         * The convertToUtcClientTime method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         *
         * @param input - The local time value to convert.
         * @returns A Date object with the time expressed in UTC.
         */
        convertToUtcClientTime(input: LocalClientTime): Date;
        /**
         * Displays an existing calendar appointment.
         *
         * The displayAppointmentForm method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.
         *
         * In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.
         *
         * In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.
         *
         * If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         *
         * @param itemId - The Exchange Web Services (EWS) identifier for an existing calendar appointment.
         */
        displayAppointmentForm(itemId: string): void;
        /**
         * Displays an existing message.
         *
         * The displayMessageForm method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.
         *
         * In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.
         *
         * If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.
         *
         * Do not use the displayMessageForm with an itemId that represents an appointment. Use the displayAppointmentForm method to display an existing appointment, and displayNewAppointmentForm to display a form to create a new appointment.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         *
         * @param itemId - The Exchange Web Services (EWS) identifier for an existing message.
         */
        displayMessageForm(itemId: string): void;
        /**
         * Displays a form for creating a new calendar appointment.
         *
         * The displayNewAppointmentForm method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.
         *
         * In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a Save button. If you have specified attendees, the form would include the attendees and a Send button.
         *
         * In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the requiredAttendees, optionalAttendees, or resources parameter, this method displays a meeting form with a Send button. If you don't specify any recipients, this method displays an appointment form with a Save & Close button.
         *
         * If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.
         *
         * Note: This method is not supported in Outlook for iOS or Outlook for Android.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
         *
         * @param parameters - An AppointmentForm describing the new appointment. All properties are optional.
         */
        displayNewAppointmentForm(parameters: AppointmentForm): void;
        /**
         * Gets a string that contains a token used to get an attachment or item from an Exchange Server.
         *
         * The getCallbackTokenAsync method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.
         *
         * You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) GetAttachment or GetItem operation to return an attachment or item. For example, you can create a remote service to get attachments from the selected item.
         *
         * Your app must have the ReadItem permission specified in its manifest to call the getCallbackTokenAsync method in read mode.
         *
         * In compose mode you must call the saveAsync method to get an item identifier to pass to the getCallbackTokenAsync method. Your app must have ReadWriteItem permissions to call the saveAsync method.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose and read</td></tr></table>
         *
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. The token is provided as a string in the asyncResult.value property.
         * @param userContext - Optional. Any state data that is passed to the asynchronous method.
         */
        getCallbackTokenAsync(callback: (result: AsyncResult) => void, userContext?: any): void;
        /**
         * Gets a token identifying the user and the Office Add-in.
         *
         * The token is provided as a string in the asyncResult.value property.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose and read</td></tr></table>
         *
         * The getUserIdentityTokenAsync method returns a token that you can use to identify and {@link https://docs.microsoft.com/outlook/add-ins/authentication | authenticate the add-in and user with a third-party system}.
         *
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         * @param userContext - Optional. Any state data that is passed to the asynchronous method.|
         */
        getUserIdentityTokenAsync(callback: (result: AsyncResult) => void, userContext?: any): void;
        /**
         * Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user's mailbox.
         *
         * In these cases, add-ins should use REST APIs to access the user's mailbox instead.
         *
         * The makeEwsRequestAsync method sends an EWS request on behalf of the add-in to Exchange.
         *
         * You cannot request Folder Associated Items with the makeEwsRequestAsync method.
         *
         * The XML request must specify UTF-8 encoding. <?xml version="1.0" encoding="utf-8"?>
         *
         * Your add-in must have the ReadWriteMailbox permission to use the makeEwsRequestAsync method. For information about using the ReadWriteMailbox permission and the EWS operations that you can call with the makeEwsRequestAsync method, see Specify permissions for mail add-in access to the user's mailbox.
         *
         * The XML result of the EWS call is provided as a string in the asyncResult.value property. If the result exceeds 1 MB in size, an error message is returned instead.
         *
         * Note: This method is not supported in the following scenarios. - In Outlook for iOS or Outlook for Android - When the add-in is loaded in a Gmail mailbox
         *
         * Note: The server administrator must set OAuthAuthentication to true on the Client Access Server EWS directory to enable the makeEwsRequestAsync method to make EWS requests.
         *
         * *Version differences*
         *
         * When you use the makeEwsRequestAsync method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to ISO-8859-1.
         *
         * `<?xml version="1.0" encoding="iso-8859-1"?>`
         *
         * You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         *
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteMailbox</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose and read</td></tr></table>
         *
         * @param data - The EWS request.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         * @param userContext - Optional. Any state data that is passed to the asynchronous method.
         */
        makeEwsRequestAsync(data: any, callback: (result: AsyncResult) => void, userContext?: any): void;
    }

    /**
     * Represents a suggested meeting found in an item. Read mode only.
     *
     * The list of meetings suggested in an email message is returned in the meetingSuggestions property of the Entities object that is returned when the getEntities or getEntitiesByType method is called on the active item.
     *
     * The start and end values are string representations of a Date object that contains the date and time at which the suggested meeting is to begin and end. The values are in the default time zone specified for the current user.
     *
     * [Api set: Mailbox 1.0]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
     */
    export interface MeetingSuggestion {
        /**
         * Gets the attendees for a suggested meeting.
         */
        attendees: EmailUser[];
        /**
         * Gets the date and time that a suggested meeting is to end.
         */
        end: string;
        /**
         * Gets the location of a suggested meeting.
         */
        location: string;
        /**
         * Gets a string that was identified as a meeting suggestion.
         */
        meetingstring: string;
        /**
         * Gets the date and time that a suggested meeting is to begin.
         */
        start: string;
        /**
         * Gets the subject of a suggested meeting.
         */
        subject: string;
    }
    /**
     * An array of NotificationMessageDetails objects are returned by the NotificationMessages.getAllAsync method.
     *
     * [Api set: Mailbox 1.3]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
     */
    export interface NotificationMessageDetails {
        /**
         * The identifier for the notification message.
         */
        key?: string;
        /**
         * Specifies the ItemNotificationMessageType of message. If type is ProgressIndicator or ErrorMessage, an icon is automatically supplied and the message is not persistent. Therefore the icon and persistent properties are not valid for these types of messages. Including them will result in an ArgumentException. If type is ProgressIndicator, the developer should remove or replace the progress indicator when the action is complete.
         */
        type: Office.MailboxEnums.ItemNotificationMessageType;
        /**
         * A reference to an icon that is defined in the manifest in the Resources section. It appears in the infobar area. It is only applicable if the type is InformationalMessage. Specifying this parameter for an unsupported type results in an exception.
         */
        icon?: string;
        /**
         * The text of the notification message. Maximum length is 150 characters. If the developer passes in a longer string, an ArgumentOutOfRange exception is thrown.
         */
        message: string;
        /**
         * Only applicable when type is InformationalMessage. If true, the message remains until removed by this add-in or dismissed by the user. If false, it is removed when the user navigates to a different item. For error notifications, the message persists until the user sees it once. Specifying this parameter for an unsupported type throws an exception.
         */
        persistent?: Boolean;
    }
    /**
     * The NotificationMessages object is returned as the notificationMessages property of an item.
     *
     * [Api set: Mailbox 1.3]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
     */
    export interface NotificationMessages {
        /**
         * Adds a notification to an item.
         *
         * There are a maximum of 5 notifications per message. Setting more will return a NumberOfNotificationMessagesExceeded error.
         *
         * @param key - A developer-specified key used to reference this notification message. Developers can use it to modify this message later. It can't be longer than 32 characters.
         * @param JSONmessage - A JSON object that contains the notification message to be added to the item. It contains a NotificationMessageDetails object.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         * 
         * In addition to this signature, the method also has the following signatures:
         * 
         * `addAsync(key: string, JSONmessage: NotificationMessageDetails): void;`
         * 
         * `addAsync(key: string, JSONmessage: NotificationMessageDetails, options: Office.AsyncContextOptions): void;`
         * 
         * `addAsync(key: string, JSONmessage: NotificationMessageDetails, callback: (result: AsyncResult) => void): void;`
         * 
         */
        addAsync(key: string, JSONmessage: NotificationMessageDetails, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Adds a notification to an item.
         *
         * There are a maximum of 5 notifications per message. Setting more will return a NumberOfNotificationMessagesExceeded error.
         *
         * @param key - A developer-specified key used to reference this notification message. Developers can use it to modify this message later. It can't be longer than 32 characters.
         * @param JSONmessage - A JSON object that contains the notification message to be added to the item. It contains a NotificationMessageDetails object.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        addAsync(key: string, JSONmessage: NotificationMessageDetails): void;
        /**
         * Adds a notification to an item.
         *
         * There are a maximum of 5 notifications per message. Setting more will return a NumberOfNotificationMessagesExceeded error.
         *
         * @param key - A developer-specified key used to reference this notification message. Developers can use it to modify this message later. It can't be longer than 32 characters.
         * @param JSONmessage - A JSON object that contains the notification message to be added to the item. It contains a NotificationMessageDetails object.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        addAsync(key: string, JSONmessage: NotificationMessageDetails, options: Office.AsyncContextOptions): void;
        /**
         * Adds a notification to an item.
         *
         * There are a maximum of 5 notifications per message. Setting more will return a NumberOfNotificationMessagesExceeded error.
         *
         * @param key - A developer-specified key used to reference this notification message. Developers can use it to modify this message later. It can't be longer than 32 characters.
         * @param JSONmessage - A JSON object that contains the notification message to be added to the item. It contains a NotificationMessageDetails object.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        addAsync(key: string, JSONmessage: NotificationMessageDetails, callback: (result: AsyncResult) => void): void;
        /**
         * Returns all keys and messages for an item.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         * 
         * In addition to the main signature, this method also has this signature:
         * 
         * `getAllAsync(callback: (result: AsyncResult) => void): void;`
         *
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         */
        getAllAsync(options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Returns all keys and messages for an item.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         *
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         */
        getAllAsync(callback: (result: AsyncResult) => void): void;
        /**
         * Removes a notification message for an item.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         * 
         * In addition to the main signature, this method also has these signatures:
         * 
         * `removeAsync(key: string): void;`
         * 
         * `removeAsync(key: string, options: Office.AsyncContextOptions): void;`
         * 
         * `removeAsync(key: string, callback: (result: AsyncResult) => void): void;`
         *
         * @param key - The key for the notification message to remove.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         */
        removeAsync(key: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Removes a notification message for an item.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         *
         * @param key - The key for the notification message to remove.
         */
        removeAsync(key: string): void;
        /**
         * Removes a notification message for an item.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         *
         * @param key - The key for the notification message to remove.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         */
        removeAsync(key: string, options: Office.AsyncContextOptions): void;
        /**
         * Removes a notification message for an item.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         *
         * @param key - The key for the notification message to remove.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         */
        removeAsync(key: string, callback: (result: AsyncResult) => void): void;
        /**
         * Replaces a notification message that has a given key with another message.
         *
         * If a notification message with the specified key doesn't exist, replaceAsync will add the notification.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         * 
         * In addition to the main signature, this method also has these signatures:
         * 
         * `replaceAsync(key: string, JSONmessage: NotificationMessageDetails): void;`
         * 
         * `replaceAsync(key: string, JSONmessage: NotificationMessageDetails, options: Office.AsyncContextOptions): void;`
         * 
         * `replaceAsync(key: string, JSONmessage: NotificationMessageDetails, callback: (result: AsyncResult) => void): void;`
         *
         * @param key - The key for the notification message to replace. It can't be longer than 32 characters.
         * @param JSONmessage - A JSON object that contains the new notification message to replace the existing message. It contains a NotificationMessageDetails object.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         */
        replaceAsync(key: string, JSONmessage: NotificationMessageDetails, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Replaces a notification message that has a given key with another message.
         *
         * If a notification message with the specified key doesn't exist, replaceAsync will add the notification.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         *
         * @param key - The key for the notification message to replace. It can't be longer than 32 characters.
         * @param JSONmessage - A JSON object that contains the new notification message to replace the existing message. It contains a NotificationMessageDetails object.
         */
        replaceAsync(key: string, JSONmessage: NotificationMessageDetails): void;
        /**
         * Replaces a notification message that has a given key with another message.
         *
         * If a notification message with the specified key doesn't exist, replaceAsync will add the notification.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         *
         * @param key - The key for the notification message to replace. It can't be longer than 32 characters.
         * @param JSONmessage - A JSON object that contains the new notification message to replace the existing message. It contains a NotificationMessageDetails object.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         */
        replaceAsync(key: string, JSONmessage: NotificationMessageDetails, options: Office.AsyncContextOptions): void;
        /**
         * Replaces a notification message that has a given key with another message.
         *
         * If a notification message with the specified key doesn't exist, replaceAsync will add the notification.
         *
         * [Api set: Mailbox 1.3]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         *
         * @param key - The key for the notification message to replace. It can't be longer than 32 characters.
         * @param JSONmessage - A JSON object that contains the new notification message to replace the existing message. It contains a NotificationMessageDetails object.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         */
        replaceAsync(key: string, JSONmessage: NotificationMessageDetails, callback: (result: AsyncResult) => void): void;
    }
    /**
     * Represents a phone number identified in an item. Read mode only.
     *
     * An array of PhoneNumber objects containing the phone numbers found in an email message is returned in the phoneNumbers property of the Entities object that is returned when you call the getEntities method on the selected item.
     *
     * [Api set: Mailbox 1.0]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
     */
    export interface PhoneNumber {
        /**
         * Gets a string containing a phone number. This string contains only the digits of the telephone number and excludes characters like parentheses and hyphens, if they exist in the original item.
         */
        phoneString: string;
        /**
         * Gets the text that was identified in an item as a phone number.
         */
        originalPhoneString: string;
        /**
         * Gets a string that identifies the type of phone number: Home, Work, Mobile, Unspecified.
         */
        type: string;
    }
    /**
     * [Api set: Mailbox 1.1]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr></table>
     */
    export interface Recipients {
        /**
         * Adds a recipient list to the existing recipients for an appointment or message.
         *
         * The recipients parameter can be an array of one of the following:
         *
         * - Strings containing SMTP email addresses
         *
         * - EmailUser objects
         *
         * - EmailAddressDetails objects
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>NumberOfRecipientsExceeded - The number of recipients exceeded 100 entries.</td></tr></table>
         * 
         * In addition to the main signature, this method also has these signatures:
         * 
         * `addAsync(recipients: (string | EmailUser | EmailAddressDetails)[]): void;`
         * 
         * `addAsync(recipients: (string | EmailUser | EmailAddressDetails)[], options: Office.AsyncContextOptions): void;`
         * 
         * `addAsync(recipients: (string | EmailUser | EmailAddressDetails)[], callback: (result: AsyncResult) => void): void;`
         *
         * @param recipients - The recipients to add to the recipients list.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If adding the recipients fails, the asyncResult.error property will contain an error code.
         */
        addAsync(recipients: (string | EmailUser | EmailAddressDetails)[], options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Adds a recipient list to the existing recipients for an appointment or message.
         *
         * The recipients parameter can be an array of one of the following:
         *
         * - Strings containing SMTP email addresses
         *
         * - EmailUser objects
         *
         * - EmailAddressDetails objects
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>NumberOfRecipientsExceeded - The number of recipients exceeded 100 entries.</td></tr></table>
         *
         * @param recipients - The recipients to add to the recipients list.
         */
        addAsync(recipients: (string | EmailUser | EmailAddressDetails)[]): void;
        /**
         * Adds a recipient list to the existing recipients for an appointment or message.
         *
         * The recipients parameter can be an array of one of the following:
         *
         * - Strings containing SMTP email addresses
         *
         * - {@link Office.EmailUser} objects
         *
         * - {@link Office.EmailAddressDetails} objects
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>NumberOfRecipientsExceeded - The number of recipients exceeded 100 entries.</td></tr></table>
         *
         * @param recipients - The recipients to add to the recipients list.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         */
        addAsync(recipients: (string | EmailUser | EmailAddressDetails)[], options: Office.AsyncContextOptions): void;
        /**
         * Adds a recipient list to the existing recipients for an appointment or message.
         *
         * The recipients parameter can be an array of one of the following:
         *
         * - Strings containing SMTP email addresses
         *
         * - {@link Office.EmailUser} objects
         *
         * - {@link Office.EmailAddressDetails} objects
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>NumberOfRecipientsExceeded - The number of recipients exceeded 100 entries.</td></tr></table>
         *
         * @param recipients - The recipients to add to the recipients list.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If adding the recipients fails, the asyncResult.error property will contain an error code.
         */
        addAsync(recipients: (string | EmailUser | EmailAddressDetails)[], callback: (result: AsyncResult) => void): void;
        /**
         * Gets a recipient list for an appointment or message.
         *
         * When the call completes, the asyncResult.value property will contain an array of{@link Office.EmailAddressDetails} objects.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr></table>
         * 
         * In addition to the main signature, this method also has this signature:
         * 
         * `getAsync(callback: (result: AsyncResult) => void): void;`
         *
         * @param options - An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         */
        getAsync(options: Office.AsyncContextOptions, callback: (result: AsyncResult) => void): void;
        /**
         * Gets a recipient list for an appointment or message.
         *
         * When the call completes, the asyncResult.value property will contain an array of {@link Office.EmailAddressDetails} objects.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr></table>
         *
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         */
        getAsync(callback: (result: AsyncResult) => void): void;
        /**
         * Sets a recipient list for an appointment or message.
         *
         * The setAsync method overwrites the current recipient list.
         *
         * The recipients parameter can be an array of one of the following:
         *
         * - Strings containing SMTP email addresses
         *
         * - {@link Office.EmailUser} objects
         *
         * - {@link Office.EmailAddressDetails} objects
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>NumberOfRecipientsExceeded - The number of recipients exceeded 100 entries.</td></tr></table>
         * 
         * In addition to the main signature, this method also has these signatures:
         * 
         * `setAsync(recipients: (string | EmailUser | EmailAddressDetails)[]): void;`
         * 
         * `setAsync(recipients: (string | EmailUser | EmailAddressDetails)[], options: Office.AsyncContextOptions): void;`
         * 
         * `setAsync(recipients: (string | EmailUser | EmailAddressDetails)[], callback: (result: AsyncResult) => void): void;`
         *
         * @param recipients - The recipients to add to the recipients list.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.  If setting the recipients fails the asyncResult.error property will contain a code that indicates any error that occurred while adding the data.
         */
        setAsync(recipients: (string | EmailUser | EmailAddressDetails)[], options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Sets a recipient list for an appointment or message.
         *
         * The setAsync method overwrites the current recipient list.
         *
         * The recipients parameter can be an array of one of the following:
         *
         * - Strings containing SMTP email addresses
         *
         * - {@link Office.EmailUser} objects
         *
         * - {@link Office.EmailAddressDetails} objects
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>NumberOfRecipientsExceeded - The number of recipients exceeded 100 entries.</td></tr></table>
         *
         * @param recipients - The recipients to add to the recipients list.
         */
        setAsync(recipients: (string | EmailUser | EmailAddressDetails)[]): void;
        /**
         * Sets a recipient list for an appointment or message.
         *
         * The setAsync method overwrites the current recipient list.
         *
         * The recipients parameter can be an array of one of the following:
         *
         * - Strings containing SMTP email addresses
         *
         * - {@link Office.EmailUser} objects
         *
         * - {@link Office.EmailAddressDetails} objects
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>NumberOfRecipientsExceeded - The number of recipients exceeded 100 entries.</td></tr></table>
         *
         * @param recipients - The recipients to add to the recipients list.
         * @param options - Optional. An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         */
        setAsync(recipients: (string | EmailUser | EmailAddressDetails)[], options: Office.AsyncContextOptions): void;
        /**
         * Sets a recipient list for an appointment or message.
         *
         * The setAsync method overwrites the current recipient list.
         *
         * The recipients parameter can be an array of one of the following:
         *
         * - Strings containing SMTP email addresses
         *
         * - {@link Office.EmailUser} objects
         *
         * - {@link Office.EmailAddressDetails} objects
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>NumberOfRecipientsExceeded - The number of recipients exceeded 100 entries.</td></tr></table>
         *
         * @param recipients - The recipients to add to the recipients list.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.  If setting the recipients fails the asyncResult.error property will contain a code that indicates any error that occurred while adding the data.
         */
        setAsync(recipients: (string | EmailUser | EmailAddressDetails)[], callback: (result: AsyncResult) => void): void;

    }

    /**
     * A file or item attachment. Used when displaying a reply form.
     */
    export interface ReplyFormAttachment {
        /**
         * Indicates the type of attachment. Must be file for a file attachment or item for an item attachment.
         */
        type: string;
        /**
         * A string that contains the name of the attachment, up to 255 characters in length.
         */
        name: string;
        /**
         * Only used if type is set to file. The URI of the location for the file.
         */
        url?: string;
        /**
         * Only used if type is set to file. If true, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.
         */
        inLine?: boolean;
        /**
         * Only used if type is set to item. The EWS item id of the attachment. This is a string up to 100 characters.
         */
        itemId?: string;
    }

    /**
     * A ReplyFormData object that contains body or attachment data and a callback function. Used when displaying a reply form.
     */
    export interface ReplyFormData {
        /**
         * A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.
         */
        htmlBody?: string;
        /**
         * An array of {@link Office.ReplyFormAttachment} that are either file or item attachments.
         */
        attachments?: ReplyFormAttachment[];
        /**
         * When the reply display call completes, the function passed in the callback parameter is called with a single parameter, asyncResult, which is an AsyncResult object.
         */
        callback?: (result: AsyncResult) => void;
    }
    /**
     * The settings created by using the methods of the RoamingSettings object are saved per add-in and per user. That is, they are available only to the add-in that created them, and only from the user's mail box in which they are saved.
     *
     * While the Outlook add-in API limits access to these settings to only the add-in that created them, these settings should not be considered secure storage. They can be accessed by Exchange Web Services or Extended MAPI. They should not be used to store sensitive information such as user credentials or security tokens.
     *
     * The name of a setting is a String, while the value can be a String, Number, Boolean, null, Object, or Array.
     *
     * The RoamingSettings object is accessible via the roamingSettings property in the Office.context namespace.
     *
     * Important: The RoamingSettings object is initialized from the persisted storage only when the add-in is first loaded. For task panes, this means that it is only initialized when the task pane first opens. If the task pane navigates to another page or reloads the current page, the in-memory object is reset to its initial values, even if your add-in has persisted changes. The persisted changes will not be available until the task pane is closed and reopened.
     *
     * [Api set: Mailbox 1.0]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>Restricted</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
     */
    export interface RoamingSettings {
        /**
         * Retrieves the specified setting.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>Restricted</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         *
         * @param name - The case-sensitive name of the setting to retrieve.
         * @returns Type: String | Number | Boolean | Object | Array
         */
        get(name: string): any;
        /**
         * Removes the specified setting
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>Restricted</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         *
         * @param name - The case-sensitive name of the setting to remove.
         */
        remove(name: string): void;
        /**
         * Saves the settings.
         *
         * Any settings previously saved by an add-in are loaded when it is initialized, so during the lifetime of the session you can just use the set and get methods to work with the in-memory copy of the settings property bag. When you want to persist the settings so that they are available the next time the add-in is used, use the saveAsync method.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>Restricted</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         *
         * @param callback - Optional? When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         */
        saveAsync(callback?: (result: AsyncResult) => void): void;
        /**
         * Sets or creates the specified setting.
         *
         * The set method creates a new setting of the specified name if it does not already exist, or sets an existing setting of the specified name. The value is stored in the document as the serialized JSON representation of its data type.
         *
         * A maximum of 32KB is available for the settings of each add-in.
         *
         * Any changes made to settings using the set function will not be saved to the server until the saveAsync function is called.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>Restricted</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         *
         * @param name - The case-sensitive name of the setting to set or create.
         * @param value - Specifies the value to be stored.
         */
        set(name: string, value: any): void;
    }
    /**
     * Provides methods to get and set the subject of an appointment or message in an Outlook add-in.
     *
     * [Api set: Mailbox 1.1]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr></table>
     */
    export interface Subject {
        /**
         * Gets the subject of an appointment or message.
         *
         * The getAsync method starts an asynchronous call to the Exchange server to get the subject of an appointment or message.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr></table>
         * 
         * In addition to the main signature, this method also has this signature:
         * 
         * `getAsync(callback: (result: AsyncResult) => void): void;`
         *
         * @param options - An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         */
        getAsync(options: Office.AsyncContextOptions, callback: (result: AsyncResult) => void): void;
        /**
         * Gets the subject of an appointment or message.
         * The getAsync method starts an asynchronous call to the Exchange server to get the subject of an appointment or message.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr></table>
         *
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         */
        getAsync(callback: (result: AsyncResult) => void): void;
        /**
         * Sets the subject of an appointment or message.
         *
         * The setAsync method starts an asynchronous call to the Exchange server to set the subject of an appointment or message. Setting the subject overwrites the current subject, but leaves any prefixes, such as "Fwd:" or "Re:" in place.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>DataExceedsMaximumSize - The subject parameter is longer than 255 characters.</td></tr></table>
         * 
         * In addition to the main signature, this method also has these signatures:
         * 
         * `setAsync(subject: string): void;`
         * 
         * `setAsync(subject: string, options: Office.AsyncContextOptions): void;`
         * 
         * `setAsync(subject: string, callback: (result: AsyncResult) => void): void;`
         *
         * @param subject - The subject of the appointment or message. The string is limited to 255 characters.
         * @param options - An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If setting the subject fails, the asyncResult.error property will contain an error code.
         */
        setAsync(subject: string, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Sets the subject of an appointment or message.
         *
         * The setAsync method starts an asynchronous call to the Exchange server to set the subject of an appointment or message. Setting the subject overwrites the current subject, but leaves any prefixes, such as "Fwd:" or "Re:" in place.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>DataExceedsMaximumSize - The subject parameter is longer than 255 characters.</td></tr></table>
         *
         * @param subject - The subject of the appointment or message. The string is limited to 255 characters.
         */
        setAsync(data: string): void;
        /**
         * Sets the subject of an appointment or message.
         *
         * The setAsync method starts an asynchronous call to the Exchange server to set the subject of an appointment or message. Setting the subject overwrites the current subject, but leaves any prefixes, such as "Fwd:" or "Re:" in place.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>DataExceedsMaximumSize - The subject parameter is longer than 255 characters.</td></tr></table>
         *
         * @param subject - The subject of the appointment or message. The string is limited to 255 characters.
         * @param options - An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         */
        setAsync(data: string, options: Office.AsyncContextOptions): void;
        /**
         * Sets the subject of an appointment or message.
         *
         * The setAsync method starts an asynchronous call to the Exchange server to set the subject of an appointment or message. Setting the subject overwrites the current subject, but leaves any prefixes, such as "Fwd:" or "Re:" in place.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>DataExceedsMaximumSize - The subject parameter is longer than 255 characters.</td></tr></table>
         *
         * @param subject - The subject of the appointment or message. The string is limited to 255 characters.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If setting the subject fails, the asyncResult.error property will contain an error code.
         */
        setAsync(data: string, callback: (result: AsyncResult) => void): void;

    }
    /**
     * Represents a suggested task identified in an item. Read mode only.
     *
     * The list of tasks suggested in an email message is returned in the taskSuggestions property of the [Entities]Entities object that is returned when the getEntities or getEntitiesByType method is called on the active item.
     *
     * [Api set: Mailbox 1.0]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Read</td></tr></table>
     */
    export interface TaskSuggestion {
        /**
         * Gets the users that should be assigned a suggested task.
         */
        assignees: EmailUser[];
        /**
         * Gets the text of an item that was identified as a task suggestion.
         */
        taskString: string;
    }
    /**
     * The Time object is returned as the start or end property of an appointment in compose mode.
     *
     * [Api set: Mailbox 1.1]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr></table>
     */
    export interface Time {
        /**
         * Gets the start or end time of an appointment.
         *
         * The date and time is provided as a Date object in the asyncResult.value property. The value is in Coordinated Universal Time (UTC). You can convert the UTC time to the local client time by using the convertToLocalClientTime method.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr></table>
         * 
         * In addition to the main signature, this method also has this signature:
         * 
         * `getAsync(callback: (result: AsyncResult) => void): void;`
         *
         * @param options - An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         */
        getAsync(options: Office.AsyncContextOptions, callback: (result: AsyncResult) => void): void;
        /**
         * Gets the start or end time of an appointment.
         *
         * The date and time is provided as a Date object in the asyncResult.value property. The value is in Coordinated Universal Time (UTC). You can convert the UTC time to the local client time by using the convertToLocalClientTime method.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr></table>
         *
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult.
         */
        getAsync(callback: (result: AsyncResult) => void): void;
        /**
         * Sets the start or end time of an appointment.
         *
         * If the setAsync method is called on the start property, the end property will be adjusted to maintain the duration of the appointment as previously set. If the setAsync method is called on the end property, the duration of the appointment will be extended to the new end time.
         *
         * The time must be in UTC; you can get the correct UTC time by using the convertToUtcClientTime method.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidEndTime - The appointment end time is before the appointment start time.</td></tr></table>
         * 
         * In addition to the main signature, this method also has these signatures:
         * 
         * `setAsync(dateTime: Date): void;`
         * 
         * `setAsync(dateTime: Date, options: Office.AsyncContextOptions): void;`
         * 
         * `setAsync(dateTime: Date, callback: (result: AsyncResult) => void): void;`
         *
         * @param dateTime - A date-time object in Coordinated Universal Time (UTC).
         * @param options - An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If setting the date and time fails, the asyncResult.error property will contain an error code.
         */
        setAsync(dateTime: Date, options?: Office.AsyncContextOptions, callback?: (result: AsyncResult) => void): void;
        /**
         * Sets the start or end time of an appointment.
         *
         * If the setAsync method is called on the start property, the end property will be adjusted to maintain the duration of the appointment as previously set. If the setAsync method is called on the end property, the duration of the appointment will be extended to the new end time.
         *
         * The time must be in UTC; you can get the correct UTC time by using the convertToUtcClientTime method.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidEndTime - The appointment end time is before the appointment start time.</td></tr></table>
         *
         * @param dateTime - A date-time object in Coordinated Universal Time (UTC).
         */
        setAsync(dateTime: Date): void;
        /**
         * Sets the start or end time of an appointment.
         *
         * If the setAsync method is called on the start property, the end property will be adjusted to maintain the duration of the appointment as previously set. If the setAsync method is called on the end property, the duration of the appointment will be extended to the new end time.
         *
         * The time must be in UTC; you can get the correct UTC time by using the convertToUtcClientTime method.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidEndTime - The appointment end time is before the appointment start time.</td></tr></table>
         *
         * @param dateTime - A date-time object in Coordinated Universal Time (UTC).
         * @param options - An object literal that contains one or more of the following properties.
         *        asyncContext: Developers can provide any object they wish to access in the callback method.
         */
        setAsync(dateTime: Date, options: Office.AsyncContextOptions): void;
        /**
         * Sets the start or end time of an appointment.
         *
         * If the setAsync method is called on the start property, the end property will be adjusted to maintain the duration of the appointment as previously set. If the setAsync method is called on the end property, the duration of the appointment will be extended to the new end time.
         *
         * The time must be in UTC; you can get the correct UTC time by using the convertToUtcClientTime method.
         *
         * [Api set: Mailbox 1.1]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadWriteItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose</td></tr>
         *
         * <tr><td>Errors</td><td>InvalidEndTime - The appointment end time is before the appointment start time.</td></tr></table>
         *
         * @param dateTime - A date-time object in Coordinated Universal Time (UTC).
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of type AsyncResult. If setting the date and time fails, the asyncResult.error property will contain an error code.
         */
        setAsync(dateTime: Date, callback: (result: AsyncResult) => void): void;

    }
    /**
     * Information about the user associated with the mailbox. This includes their account type, display name, email adddress, and time zone.
     * 
     * [Api set: Mailbox 1.0]
     *
     * @remarks
     * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
     *
     * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
     */
    export interface UserProfile {
        /**
         * Gets the user's display name.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        displayName: string;
        /**
         * Gets the user's display name.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        emailAddress: string;
        /**
         * Gets the user's SMTP email address.
         *
         * [Api set: Mailbox 1.0]
         *
         * @remarks
         * <table><tr><td>{@link https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions | Minimum permission level}</td><td>ReadItem</td></tr>
         *
         * <tr><td>{@link https://docs.microsoft.com/outlook/add-ins/#extension-points | Applicable Outlook mode}</td><td>Compose or read</td></tr></table>
         */
        timeZone: string;
    }
}


////////////////////////////////////////////////////////////////
/////////////////////// End Exchange APIs //////////////////////
////////////////////////////////////////////////////////////////