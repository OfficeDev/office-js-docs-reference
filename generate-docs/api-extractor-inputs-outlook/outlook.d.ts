import {Office as CommonAPI} from "../api-extractor-inputs-office/office"
////////////////////////////////////////////////////////////////
////////////////////// Begin Exchange APIs /////////////////////
////////////////////////////////////////////////////////////////

export declare namespace Office {
    export namespace MailboxEnums {
        /**
         * Specifies the type of custom action in a notification message.
         *
         * @remarks
         * [Api set: Mailbox 1.10]
         */
        enum ActionType {
            /**
             * The `showTaskPane` action.
             */
            ShowTaskPane = "showTaskPane"
        }
        /**
         * Specifies the {@link Office.Sensitivity | sensitivity level} of an appointment.
         *
         * @remarks
         * [Api set: Mailbox preview]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @beta
         */
        enum AppointmentSensitivityType {
            /**
             * The item needs no special treatment.
             *
             * @remarks
             * [Api set: Mailbox preview]
             */
            Normal = "normal",
            /**
             * Treat the item as personal.
             *
             * @remarks
             * [Api set: Mailbox preview]
             */
            Personal = "personal",
            /**
             * Treat the item as private.
             *
             * @remarks
             * [Api set: Mailbox preview]
             */
            Private = "private",
            /**
             * Treat the item as confidential.
             *
             * @remarks
             * [Api set: Mailbox preview]
             */
            Confidential = "confidential"
        }
        /**
         * Specifies the formatting that applies to an attachment's content.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        enum AttachmentContentFormat {
            /**
             * The content of the attachment is returned as a base64-encoded string.
             */
            Base64 = "base64",
            /**
             * The content of the attachment is returned as a string representing a URL.
             */
            Url = "url",
            /**
             * The content of the attachment is returned as a string representing an .eml formatted file.
             */
            Eml = "eml",
            /**
             * The content of the attachment is returned as a string representing an .icalendar formatted file.
             */
            ICalendar = "iCalendar"
        }
        /**
         * Specifies whether an attachment was added to or removed from an item.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        enum AttachmentStatus {
            /**
             * An attachment was added to the item.
             */
            Added = "added",
            /**
             * An attachment was removed from the item.
             */
            Removed = "removed"
        }
        /**
         * Specifies an attachment's type.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        enum AttachmentType {
            /**
             * The attachment is a file.
             */
            File = "file",
            /**
             * The attachment is an Exchange item.
             */
            Item = "item",
            /**
             * The attachment is stored in a cloud location, such as OneDrive.
             *
             * **Important**: In Read mode, the `id` property of the attachment's {@link Office.AttachmentDetails | details} object
             * contains a URL to the file.
             * From requirement set 1.8, the `url` property included in the attachment's {@link Office.AttachmentDetailsCompose | details} object
             * contains a URL to the file in Compose mode.
             */
            Cloud = "cloud"
        }
        /**
         * Specifies the category color.
         *
         * **Note**: The actual color depends on how the Outlook client renders it.
         * In this case, the colors noted on each preset are for the Outlook desktop client.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        enum CategoryColor {
            /**
             * Default color or no color mapped.
             */
            None,
            /**
             * Red
             */
            Preset0,
            /**
             * Orange
             */
            Preset1,
            /**
             * Brown
             */
            Preset2,
            /**
             * Yellow
             */
            Preset3,
            /**
             * Green
             */
            Preset4,
            /**
             * Teal
             */
            Preset5,
            /**
             * Olive
             */
            Preset6,
            /**
             * Blue
             */
            Preset7,
            /**
             * Purple
             */
            Preset8,
            /**
             * Cranberry
             */
            Preset9,
            /**
             * Steel
             */
            Preset10,
            /**
             * DarkSteel
             */
            Preset11,
            /**
             * Gray
             */
            Preset12,
            /**
             * DarkGray
             */
            Preset13,
            /**
             * Black
             */
            Preset14,
            /**
             * DarkRed
             */
            Preset15,
            /**
             * DarkOrange
             */
            Preset16,
            /**
             * DarkBrown
             */
            Preset17,
            /**
             * DarkYellow
             */
            Preset18,
            /**
             * DarkGreen
             */
            Preset19,
            /**
             * DarkTeal
             */
            Preset20,
            /**
             * DarkOlive
             */
            Preset21,
            /**
             * DarkBlue
             */
            Preset22,
            /**
             * DarkPurple
             */
            Preset23,
            /**
             * DarkCranberry
             */
            Preset24
        }
        /**
         * Specifies a message's compose type.
         *
         * @remarks
         * [Api set: Mailbox 1.10]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         */
        enum ComposeType {
            /**
             * Reply.
             */
            Reply = "reply",
            /**
             * New mail.
             */
            NewMail = "newMail",
            /**
             * Forward.
             */
            Forward = "forward"
        }
        /**
         * Specifies the day of week or type of day.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        enum Days {
            /**
             * Monday
             */
            Mon = "mon",
            /**
             * Tuesday
             */
            Tue = "tue",
            /**
             * Wednesday
             */
            Wed = "wed",
            /**
             * Thursday
             */
            Thu = "thu",
            /**
             * Friday
             */
            Fri = "fri",
            /**
             * Saturday
             */
            Sat = "sat",
            /**
             * Sunday
             */
            Sun = "sun",
            /**
             * Week day (excludes weekend days): 'Mon', 'Tue', 'Wed', 'Thu', and 'Fri'.
             */
            Weekday = "weekday",
            /**
             * Weekend day: 'Sat' and 'Sun'.
             */
            WeekendDay = "weekendDay",
            /**
             * Day of week.
             */
            Day = "day"
        }
        /**
         * This bitmask represents a delegate's permissions on a shared folder, or a user's permissions on a shared mailbox.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        enum DelegatePermissions {
            /**
             * Delegate or user has permission to read items.
             */
            Read = 1,
            /**
             * Delegate or user has permission to create and write items.
             */
            Write = 2,
            /**
             * Delegate or user has permission to delete only the items they created.
             */
            DeleteOwn = 4,
            /**
             * Delegate or user has permission to delete any items.
             */
            DeleteAll = 8,
            /**
             * Delegate or user has permission to edit only they items they created.
             */
            EditOwn = 16,
            /**
             * Delegate or user has permission to edit any items.
             */
            EditAll = 32
        }
        /**
         * Specifies an entity's type.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
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
         * Action types supported by {@link https://learn.microsoft.com/javascript/api/office/office.eventtype | Office.EventType.InfobarClicked}.
         *
         * @remarks
         * [Api set: Mailbox 1.10]
         */
        enum InfobarActionType {
            /**
             * Dismiss action was selected.
             *
             * @remarks
             * [Api set: Mailbox 1.10]
             */
            Dismiss = 1
        }
        /**
         * Type of notification allowed by {@link https://learn.microsoft.com/javascript/api/office/office.eventtype | Office.EventType.InfobarClicked}.
         *
         * @remarks
         * [Api set: Mailbox 1.10]
         */
        enum InfobarType {
            /**
             * Notification displays an informational message.
             *
             * @remarks
             * [Api set: Mailbox 1.10]
             */
            Informational = 0,
            /**
             * Notification displays a progress indicator.
             *
             * @remarks
             * [Api set: Mailbox 1.10]
             */
            ProgressIndicator = 1,
            /**
             * Notification displays an error message.
             *
             * @remarks
             * [Api set: Mailbox 1.10]
             */
            Error = 2,
            /**
             * Notification displays an informational message with actions.
             *
             * @remarks
             * [Api set: Mailbox 1.10]
             */
            Insight = 3
        }
        /**
         * Specifies the notification message type for an appointment or message.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        enum ItemNotificationMessageType {
            /**
             * The notification message is a progress indicator.
             */
            ProgressIndicator = "progressIndicator",
            /**
             * The notification message is an informational message.
             */
            InformationalMessage = "informationalMessage",
            /**
             * The notification message is an error message.
             */
            ErrorMessage = "errorMessage",
            /**
             * The notification message is an informational message with actions.
             *
             * @remarks
             * [Api set: Mailbox 1.10]
             */
            InsightMessage = "insightMessage"
        }
        /**
         * Specifies an item's type.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
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
         * Specifies an appointment location's type.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        enum LocationType {
            /**
             * A custom location. Custom locations don't have an SMTP address.
             * 
             * **Note**: {@link https://support.microsoft.com/office/88ff6c60-0a1d-4b54-8c9d-9e1a71bc3023 | Personal contact groups}
             * added as appointment locations aren't returned by the
             * {@link https://learn.microsoft.com/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1) | EnhancedLocation.getAsync} method.
             */
            Custom = "custom",
            /**
             * A conference room or similar resource that has an SMTP address.
             */
            Room = "room"
        }
        /**
         * Specifies the month.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        enum Month {
            /**
             * January
             */
            Jan = "jan",
            /**
             * February
             */
            Feb = "feb",
            /**
             * March
             */
            Mar = "mar",
            /**
             * April
             */
            Apr = "apr",
            /**
             * May
             */
            May = "may",
            /**
             * June
             */
            Jun = "jun",
            /**
             * July
             */
            Jul = "jul",
            /**
             * August
             */
            Aug = "aug",
            /**
             * September
             */
            Sep = "sep",
            /**
             * October
             */
            Oct = "oct",
            /**
             * November
             */
            Nov = "nov",
            /**
             * December
             */
            Dec = "dec"
        }
        /**
         * Represents the current view of Outlook on the web.
         */
        enum OWAView {
            /**
             * Narrow one-column view. Displayed when the screen width is less than 436 pixels.
             * For example, Outlook on the web uses this view on the entire screen of older smartphones.
             */
            OneColumnNarrow = "OneColumnNarrow",
            /**
             * One-column view. Displayed when the screen width is greater than or equal to 436 pixels,
             * but less than 536 pixels. For example, Outlook on the web uses this view on the entire screen of newer smartphones.
             */
            OneColumn = "OneColumn",
            /**
             * Two-column view. Displayed when the screen width is greater than or equal to 536 pixels,
             * but less than 780 pixels. For example, Outlook on the web uses this view on most tablets.
             */
            TwoColumns = "TwoColumns",
            /**
             * Three-column view. Displayed when the screen width is greater than or equal to 780 pixels.
             * For example, Outlook on the web uses this view in a full screen window on a desktop computer.
             */
            ThreeColumns = "ThreeColumns"
        }
        /**
         * Specifies the type of recipient of a message or appointment.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         * 
         * **Important**: A `recipientType` property value isn't returned by the 
         * {@link https://learn.microsoft.com/javascript/api/outlook/office.from?view=outlook-js-preview#outlook-office-from-getasync-member(1) | Office.context.mailbox.item.from.getAsync} 
         * and {@link https://learn.microsoft.com/javascript/api/outlook/office.organizer?view=outlook-js-preview#outlook-office-organizer-getasync-member(1) | Office.context.mailbox.item.organizer.getAsync} methods.
         * The email sender or appointment organizer is always a user whose email address is on the Exchange server.
         */
        enum RecipientType {
            /**
             * Specifies the recipient is a distribution list containing a list of email addresses.
             */
            DistributionList = "distributionList",
            /**
             * Specifies the recipient is an SMTP email address on the Exchange server.
             */
            User = "user",
            /**
             * Specifies the recipient is an SMTP email address that isn't on the Exchange server. It also refers to a recipient added from a personal Outlook address book.
             * 
             * **Note**: In Outlook on Windows starting with Version 2210 (Build 15813.20002), Global Address Book (GAL) recipients saved to a personal address book return
             * the `ExternalUser` value, even if their SMTP email address appears on the Exchange server. Recipients return a `User` value only if they're directly
             * added or resolved against the GAL.
             */
            ExternalUser = "externalUser",
            /**
             * Specifies the recipient isn't one of the other recipient types. It also refers to a recipient that isn't resolved against the Exchange address book,
             * and is therefore treated as an external SMTP address.
             */
            Other = "other"
        }
        /**
         * Specifies the time zone applied to the recurrence.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        enum RecurrenceTimeZone {
            /**
             * Afghanistan Standard Time
             */
            AfghanistanStandardTime = "Afghanistan Standard Time",
            /**
             * Alaskan Standard Time
             */
            AlaskanStandardTime = "Alaskan Standard Time",
            /**
             * Aleutian Standard Time
             */
            AleutianStandardTime = "Aleutian Standard Time",
            /**
             * Altai Standard Time
             */
            AltaiStandardTime = "Altai Standard Time",
            /**
             * Arab Standard Time
             */
            ArabStandardTime = "Arab Standard Time",
            /**
             * Arabian Standard Time
             */
            ArabianStandardTime = "Arabian Standard Time",
            /**
             * Arabic Standard Time
             */
            ArabicStandardTime = "Arabic Standard Time",
            /**
             * Argentina Standard Time
             */
            ArgentinaStandardTime = "Argentina Standard Time",
            /**
             * Astrakhan Standard Time
             */
            AstrakhanStandardTime = "Astrakhan Standard Time",
            /**
             * Atlantic Standard Time
             */
            AtlanticStandardTime = "Atlantic Standard Time",
            /**
             * Australia Central Standard Time
             */
            AUSCentralStandardTime = "AUS Central Standard Time",
            /**
             * Australia Central West Standard Time
             */
            AusCentralW_StandardTime = "Aus Central W. Standard Time",
            /**
             * AUS Eastern Standard Time
             */
            AUSEasternStandardTime = "AUS Eastern Standard Time",
            /**
             * Azerbaijan Standard Time
             */
            AzerbaijanStandardTime = "Azerbaijan Standard Time",
            /**
             * Azores Standard Time
             */
            AzoresStandardTime = "Azores Standard Time",
            /**
             * Bahia Standard Time
             */
            BahiaStandardTime = "Bahia Standard Time",
            /**
             * Bangladesh Standard Time
             */
            BangladeshStandardTime = "Bangladesh Standard Time",
            /**
             * Belarus Standard Time
             */
            BelarusStandardTime = "Belarus Standard Time",
            /**
             * Bougainville Standard Time
             */
            BougainvilleStandardTime = "Bougainville Standard Time",
            /**
             * Canada Central Standard Time
             */
            CanadaCentralStandardTime = "Canada Central Standard Time",
            /**
             * Cape Verde Standard Time
             */
            CapeVerdeStandardTime = "Cape Verde Standard Time",
            /**
             * Caucasus Standard Time
             */
            CaucasusStandardTime = "Caucasus Standard Time",
            /**
             * Central Australia Standard Time
             */
            CenAustraliaStandardTime = "Cen. Australia Standard Time",
            /**
             * Central America Standard Time
             */
            CentralAmericaStandardTime = "Central America Standard Time",
            /**
             * Central Asia Standard Time
             */
            CentralAsiaStandardTime = "Central Asia Standard Time",
            /**
             * Central Brazilian Standard Time
             */
            CentralBrazilianStandardTime = "Central Brazilian Standard Time",
            /**
             * Central Europe Standard Time
             */
            CentralEuropeStandardTime = "Central Europe Standard Time",
            /**
             * Central European Standard Time
             */
            CentralEuropeanStandardTime = "Central European Standard Time",
            /**
             * Central Pacific Standard Time
             */
            CentralPacificStandardTime = "Central Pacific Standard Time",
            /**
             * Central Standard Time
             */
            CentralStandardTime = "Central Standard Time",
            /**
             * Central Standard Time (Mexico)
             */
            CentralStandardTime_Mexico = "Central Standard Time (Mexico)",
            /**
             * Chatham Islands Standard Time
             */
            ChathamIslandsStandardTime = "Chatham Islands Standard Time",
            /**
             * China Standard Time
             */
            ChinaStandardTime = "China Standard Time",
            /**
             * Cuba Standard Time
             */
            CubaStandardTime = "Cuba Standard Time",
            /**
             * Dateline Standard Time
             */
            DatelineStandardTime = "Dateline Standard Time",
            /**
             * East Africa Standard Time
             */
            E_AfricaStandardTime = "E. Africa Standard Time",
            /**
             * East Australia Standard Time
             */
            E_AustraliaStandardTime = "E. Australia Standard Time",
            /**
             * East Europe Standard Time
             */
            E_EuropeStandardTime = "E. Europe Standard Time",
            /**
             * East South America Standard Time
             */
            E_SouthAmericaStandardTime = "E. South America Standard Time",
            /**
             * Easter Island Standard Time
             */
            EasterIslandStandardTime = "Easter Island Standard Time",
            /**
             * Eastern Standard Time
             */
            EasternStandardTime = "Eastern Standard Time",
            /**
             * Eastern Standard Time (Mexico)
             */
            EasternStandardTime_Mexico = "Eastern Standard Time (Mexico)",
            /**
             * Egypt Standard Time
             */
            EgyptStandardTime = "Egypt Standard Time",
            /**
             * Ekaterinburg Standard Time
             */
            EkaterinburgStandardTime = "Ekaterinburg Standard Time",
            /**
             * Fiji Standard Time
             */
            FijiStandardTime = "Fiji Standard Time",
            /**
             * FLE Standard Time
             */
            FLEStandardTime = "FLE Standard Time",
            /**
             * Georgian Standard Time
             */
            GeorgianStandardTime = "Georgian Standard Time",
            /**
             * GMT Standard Time
             */
            GMTStandardTime = "GMT Standard Time",
            /**
             * Greenland Standard Time
             */
            GreenlandStandardTime = "Greenland Standard Time",
            /**
             * Greenwich Standard Time
             */
            GreenwichStandardTime = "Greenwich Standard Time",
            /**
             * GTB Standard Time
             */
            GTBStandardTime = "GTB Standard Time",
            /**
             * Haiti Standard Time
             */
            HaitiStandardTime = "Haiti Standard Time",
            /**
             * Hawaiian Standard Time
             */
            HawaiianStandardTime = "Hawaiian Standard Time",
            /**
             * India Standard Time
             */
            IndiaStandardTime = "India Standard Time",
            /**
             * Iran Standard Time
             */
            IranStandardTime = "Iran Standard Time",
            /**
             * Israel Standard Time
             */
            IsraelStandardTime = "Israel Standard Time",
            /**
             * Jordan Standard Time
             */
            JordanStandardTime = "Jordan Standard Time",
            /**
             * Kaliningrad Standard Time
             */
            KaliningradStandardTime = "Kaliningrad Standard Time",
            /**
             * Kamchatka Standard Time
             */
            KamchatkaStandardTime = "Kamchatka Standard Time",
            /**
             * Korea Standard Time
             */
            KoreaStandardTime = "Korea Standard Time",
            /**
             * Libya Standard Time
             */
            LibyaStandardTime = "Libya Standard Time",
            /**
             * Line Islands Standard Time
             */
            LineIslandsStandardTime = "Line Islands Standard Time",
            /**
             * Lord Howe Standard Time
             */
            LordHoweStandardTime = "Lord Howe Standard Time",
            /**
             * Magadan Standard Time
             */
            MagadanStandardTime = "Magadan Standard Time",
            /**
             * Magallanes Standard Time
             */
            MagallanesStandardTime = "Magallanes Standard Time",
            /**
             * Marquesas Standard Time
             */
            MarquesasStandardTime = "Marquesas Standard Time",
            /**
             * Mauritius Standard Time
             */
            MauritiusStandardTime = "Mauritius Standard Time",
            /**
             * Mid-Atlantic Standard Time
             */
            MidAtlanticStandardTime = "Mid-Atlantic Standard Time",
            /**
             * Middle East Standard Time
             */
            MiddleEastStandardTime = "Middle East Standard Time",
            /**
             * Montevideo Standard Time
             */
            MontevideoStandardTime = "Montevideo Standard Time",
            /**
             * Morocco Standard Time
             */
            MoroccoStandardTime = "Morocco Standard Time",
            /**
             * Mountain Standard Time
             */
            MountainStandardTime = "Mountain Standard Time",
            /**
             * Mountain Standard Time (Mexico)
             */
            MountainStandardTime_Mexico = "Mountain Standard Time (Mexico)",
            /**
             * Myanmar Standard Time
             */
            MyanmarStandardTime = "Myanmar Standard Time",
            /**
             * North Central Asia Standard Time
             */
            N_CentralAsiaStandardTime = "N. Central Asia Standard Time",
            /**
             * Namibia Standard Time
             */
            NamibiaStandardTime = "Namibia Standard Time",
            /**
             * Nepal Standard Time
             */
            NepalStandardTime = "Nepal Standard Time",
            /**
             * New Zealand Standard Time
             */
            NewZealandStandardTime = "New Zealand Standard Time",
            /**
             * Newfoundland Standard Time
             */
            NewfoundlandStandardTime = "Newfoundland Standard Time",
            /**
             * Norfolk Standard Time
             */
            NorfolkStandardTime = "Norfolk Standard Time",
            /**
             * North Asia East Standard Time
             */
            NorthAsiaEastStandardTime = "North Asia East Standard Time",
            /**
             * North Asia Standard Time
             */
            NorthAsiaStandardTime = "North Asia Standard Time",
            /**
             * North Korea Standard Time
             */
            NorthKoreaStandardTime = "North Korea Standard Time",
            /**
             * Omsk Standard Time
             */
            OmskStandardTime = "Omsk Standard Time",
            /**
             * Pacific SA Standard Time
             */
            PacificSAStandardTime = "Pacific SA Standard Time",
            /**
             * Pacific Standard Time
             */
            PacificStandardTime = "Pacific Standard Time",
            /**
             * Pacific Standard Time (Mexico)
             */
            PacificStandardTimeMexico = "Pacific Standard Time (Mexico)",
            /**
             * Pakistan Standard Time
             */
            PakistanStandardTime = "Pakistan Standard Time",
            /**
             * Paraguay Standard Time
             */
            ParaguayStandardTime = "Paraguay Standard Time",
            /**
             * Romance Standard Time
             */
            RomanceStandardTime = "Romance Standard Time",
            /**
             * Russia Time Zone 10
             */
            RussiaTimeZone10 = "Russia Time Zone 10",
            /**
             * Russia Time Zone 11
             */
            RussiaTimeZone11 = "Russia Time Zone 11",
            /**
             * Russia Time Zone 3
             */
            RussiaTimeZone3 = "Russia Time Zone 3",
            /**
             * Russian Standard Time
             */
            RussianStandardTime = "Russian Standard Time",
            /**
             * SA Eastern Standard Time
             */
            SAEasternStandardTime = "SA Eastern Standard Time",
            /**
             * SA Pacific Standard Time
             */
            SAPacificStandardTime = "SA Pacific Standard Time",
            /**
             * SA Western Standard Time
             */
            SAWesternStandardTime = "SA Western Standard Time",
            /**
             * Saint Pierre Standard Time
             */
            SaintPierreStandardTime = "Saint Pierre Standard Time",
            /**
             * Sakhalin Standard Time
             */
            SakhalinStandardTime = "Sakhalin Standard Time",
            /**
             * Samoa Standard Time
             */
            SamoaStandardTime = "Samoa Standard Time",
            /**
             * Saratov Standard Time
             */
            SaratovStandardTime = "Saratov Standard Time",
            /**
             * Southeast Asia Standard Time
             */
            SEAsiaStandardTime = "SE Asia Standard Time",
            /**
             * Singapore Standard Time
             */
            SingaporeStandardTime = "Singapore Standard Time",
            /**
             * South Africa Standard Time
             */
            SouthAfricaStandardTime = "South Africa Standard Time",
            /**
             * Sri Lanka Standard Time
             */
            SriLankaStandardTime = "Sri Lanka Standard Time",
            /**
             * Sudan Standard Time
             */
            SudanStandardTime = "Sudan Standard Time",
            /**
             * Syria Standard Time
             */
            SyriaStandardTime = "Syria Standard Time",
            /**
             * Taipei Standard Time
             */
            TaipeiStandardTime = "Taipei Standard Time",
            /**
             * Tasmania Standard Time
             */
            TasmaniaStandardTime = "Tasmania Standard Time",
            /**
             * Tocantins Standard Time
             */
            TocantinsStandardTime = "Tocantins Standard Time",
            /**
             * Tokyo Standard Time
             */
            TokyoStandardTime = "Tokyo Standard Time",
            /**
             * Tomsk Standard Time
             */
            TomskStandardTime = "Tomsk Standard Time",
            /**
             * Tonga Standard Time
             */
            TongaStandardTime = "Tonga Standard Time",
            /**
             * Transbaikal Standard Time
             */
            TransbaikalStandardTime = "Transbaikal Standard Time",
            /**
             * Turkey Standard Time
             */
            TurkeyStandardTime = "Turkey Standard Time",
            /**
             * Turks And Caicos Standard Time
             */
            TurksAndCaicosStandardTime = "Turks And Caicos Standard Time",
            /**
             * Ulaanbaatar Standard Time
             */
            UlaanbaatarStandardTime = "Ulaanbaatar Standard Time",
            /**
             * United States Eastern Standard Time
             */
            USEasternStandardTime = "US Eastern Standard Time",
            /**
             * United States Mountain Standard Time
             */
            USMountainStandardTime = "US Mountain Standard Time",
            /**
             * Coordinated Universal Time (UTC)
             */
            UTC = "UTC",
            /**
             * Coordinated Universal Time (UTC) + 12 hours
             */
            UTCPLUS12 = "UTC+12",
            /**
             * Coordinated Universal Time (UTC) + 13 hours
             */
            UTCPLUS13 = "UTC+13",
            /**
             * Coordinated Universal Time (UTC) - 2 hours
             */
            UTCMINUS02 = "UTC-02",
            /**
             * Coordinated Universal Time (UTC) - 8 hours
             */
            UTCMINUS08 = "UTC-08",
            /**
             * Coordinated Universal Time (UTC) - 9 hours
             */
            UTCMINUS09 = "UTC-09",
            /**
             * Coordinated Universal Time (UTC) - 11 hours
             */
            UTCMINUS11 = "UTC-11",
            /**
             * Venezuela Standard Time
             */
            VenezuelaStandardTime = "Venezuela Standard Time",
            /**
             * Vladivostok Standard Time
             */
            VladivostokStandardTime = "Vladivostok Standard Time",
            /**
             * West Australia Standard Time
             */
            W_AustraliaStandardTime = "W. Australia Standard Time",
            /**
             * West Central Africa Standard Time
             */
            W_CentralAfricaStandardTime = "W. Central Africa Standard Time",
            /**
             * West Europe Standard Time
             */
            W_EuropeStandardTime = "W. Europe Standard Time",
            /**
             * West Mongolia Standard Time
             */
            W_MongoliaStandardTime = "W. Mongolia Standard Time",
            /**
             * West Asia Standard Time
             */
            WestAsiaStandardTime = "West Asia Standard Time",
            /**
             * West Bank Standard Time
             */
            WestBankStandardTime = "West Bank Standard Time",
            /**
             * West Pacific Standard Time
             */
            WestPacificStandardTime = "West Pacific Standard Time",
            /**
             * Yakutsk Standard Time
             */
            YakutskStandardTime = "Yakutsk Standard Time"
        }
        /**
         * Specifies the type of recurrence.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        enum RecurrenceType {
            /**
             * Daily.
             */
            Daily = "daily",
            /**
             * Weekday.
             */
            Weekday = "weekday",
            /**
             * Weekly.
             */
            Weekly = "weekly",
            /**
             * Monthly.
             */
            Monthly = "monthly",
            /**
             * Yearly.
             */
            Yearly = "yearly"
        }
        /**
         * Specifies the type of response to a meeting invitation.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
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
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
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
        /**
         * Specifies the source of the selected data in an item (see `Office.mailbox.item.getSelectedDataAsync` for details).
         *
         * @remarks
         * [Api set: Mailbox 1.2]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         */
        enum SourceProperty {
            /**
             * The source of the data is from the body of the item.
             */
            Body = "body",
            /**
             * The source of the data is from the subject of the item.
             */
            Subject = "subject"
        }
        /**
         * Specifies the week of the month.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        enum WeekNumber {
            /**
             * First week of the month.
             */
            First = "first",
            /**
             * Second week of the month.
             */
            Second = "second",
            /**
             * Third week of the month.
             */
            Third = "third",
            /**
             * Fourth week of the month.
             */
            Fourth = "fourth",
            /**
             * Last week of the month.
             */
            Last = "last"
        }
    }
    /**
     * Provides an option for the data format.
     */
    export interface CoercionTypeOptions {
        /**
         * The desired data format.
         */
        coercionType?: CommonAPI.CoercionType | string;
    }
    /**
     * The subclass of {@link Office.Item | Item} dealing with appointments.
     *
     * **Important**: This is an internal Outlook object, not directly exposed through existing interfaces.
     * You should treat this as a mode of `Office.context.mailbox.item`. For more information, refer to the
     * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item | Object Model} page.
     *
     * Child interfaces:
     *
     * - {@link Office.AppointmentCompose | AppointmentCompose}
     *
     * - {@link Office.AppointmentRead | AppointmentRead}
     */
    export interface Appointment extends Item {
    }
    /**
     * The appointment organizer mode of {@link Office.Item | Office.context.mailbox.item}.
     *
     * **Important**: This is an internal Outlook object, not directly exposed through existing interfaces.
     * You should treat this as a mode of `Office.context.mailbox.item`. For more information, refer to the
     * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item | Object Model} page.
     *
     * Parent interfaces:
     *
     * - {@link Office.ItemCompose | ItemCompose}
     *
     * - {@link Office.Appointment | Appointment}
     */
    export interface AppointmentCompose extends Appointment, ItemCompose {
         /**
         * Gets an object that provides methods for manipulating the body of an item.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        body: Body;
        /**
         * Gets an object that provides methods for managing the item's categories.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        categories: Categories;
        /**
         * Gets or sets the date and time that the appointment is to end.
         *
         * The `end` property is a {@link Office.Time | Time} object expressed as a Coordinated Universal Time (UTC) date and time value.
         * You can use the `convertToLocalClientTime` method to convert the `end` property value to the client's local date and time.
         *
         * When you use the `Time.setAsync` method to set the end time, you should use the `convertToUtcClientTime` method to convert the local time on
         * the client to UTC for the server.
         *
         * **Important**: In the Windows client, you can't use this property to update the end of a recurrence.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        end: Time;
        /**
         * Gets or sets the locations of the appointment. The `enhancedLocation` property returns an {@link Office.EnhancedLocation | EnhancedLocation}
         * object that provides methods to get, remove, or add locations on an item.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        enhancedLocation: EnhancedLocation;
        /**
         * Gets or sets the {@link Office.IsAllDayEvent} property of an appointment.
         *
         * @remarks
         * [Api set: Mailbox preview]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @beta
         */
        isAllDayEvent: IsAllDayEvent;
        /**
         * Gets the type of item that an instance represents.
         *
         * The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        itemType: MailboxEnums.ItemType | string;
        /**
         * Gets or sets the location of an appointment. The `location` property returns a {@link Office.Location | Location} object that provides methods that are
         * used to get and set the location of the appointment.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        location: Location;
        /**
         * Gets the notification messages for an item.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        notificationMessages: NotificationMessages;
        /**
         * Provides access to the optional attendees of an event. The type of object and level of access depend on the mode of the current item.
         *
         * The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the
         * optional attendees for a meeting. However, depending on the client/platform (i.e., Windows, Mac, etc.), limits may apply on how many
         * recipients you can get or update. See the {@link Office.Recipients | Recipients} object for more details.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        optionalAttendees: Recipients;
        /**
         * Gets the organizer for the specified meeting.
         *
         * The `organizer` property returns an {@link Office.Organizer | Organizer} object that provides a method to get the organizer value.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        organizer: Organizer;
        /**
         * Gets or sets the recurrence pattern of an appointment.
         *
         * The `recurrence` property returns a recurrence object for recurring appointments or meetings requests if an item is a series or an instance
         * in a series. `null` is returned for single appointments and meeting requests of single appointments.
         *
         * **Note**: Meeting requests have an `itemClass` value of `IPM.Schedule.Meeting.Request`.
         *
         * **Note**: If the recurrence object is null, this indicates that the object is a single appointment or a meeting request of a single
         * appointment and NOT a part of a series.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        recurrence: Recurrence;
        /**
         * Provides access to the required attendees of an event. The type of object and level of access depend on the mode of the current item.
         *
         * The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the
         * required attendees for a meeting. However, depending on the client/platform (i.e., Windows, Mac, etc.), limits may apply on how many
         * recipients you can get or update. See the {@link Office.Recipients | Recipients} object for more details.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        requiredAttendees: Recipients;
        /**
         * Gets or sets the {@link Office.Sensitivity | sensitivity level} of an appointment.
         * For information about sensitivity levels, see
         * {@link https://support.microsoft.com/office/4a76d05b-6c29-4a0d-9096-71784a6b12c1 | Mark your email as Normal, Personal, Private, or Confidential}.
         *
         * @remarks
         * [Api set: Mailbox preview]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @beta
         */
        sensitivity: Sensitivity;
        /**
         * Gets the object to get or set the {@link Office.SensitivityLabel | sensitivity label} of an appointment.
         * 
         * @remarks
         * [Api set: Mailbox preview]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * **Important**: To use the sensitivity label feature in your add-in, you must have a Microsoft 365 E5 subscription.
         *
         * To learn more about how to manage sensitivity labels in your add-in, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/sensitivity-label | Manage the sensitivity label of your message or appointment in compose mode (preview)}.
         *
         * @beta
         */
        sensitivityLabel: SensitivityLabel;
        /**
         * Gets the ID of the series that an instance belongs to.
         *
         * In Outlook on the web and desktop clients, the `seriesId` property returns the Exchange Web Services (EWS) ID of the parent (series) item
         * that this item belongs to. However, on iOS and Android, the seriesId returns the REST ID of the parent item.
         *
         * **Note**: The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.
         * The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.
         * Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`.
         * For more details, see {@link https://learn.microsoft.com/office/dev/add-ins/outlook/use-rest-api | Use the Outlook REST APIs from an Outlook add-in}.
         *
         * The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests
         * and returns `undefined` for any other items that are not meeting requests.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        seriesId: string;
        /**
         * Manages the {@link Office.SessionData | SessionData} of an item in Compose mode.
         *
         * **Important**: The entire SessionData object is limited to 50,000 characters per add-in.
         *
         * @remarks
         * [Api set: Mailbox 1.11]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        sessionData: SessionData;
        /**
         * Gets or sets the date and time that the appointment is to begin.
         *
         * The `start` property is a {@link Office.Time | Time} object expressed as a Coordinated Universal Time (UTC) date and time value.
         * You can use the `convertToLocalClientTime` method to convert the value to the client's local date and time.
         *
         * When you use the `Time.setAsync` method to set the start time, you should use the `convertToUtcClientTime` method to convert the local time on
         * the client to UTC for the server.
         *
         * **Important**: In the Windows client, you can't use this property to update the start of a recurrence.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        start: Time;
        /**
         * Gets or sets the description that appears in the subject field of an item.
         *
         * The `subject` property gets or sets the entire subject of the item, as sent by the email server.
         *
         * The `subject` property returns a `Subject` object that provides methods to get and set the subject.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        subject: Subject;

        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * You can subsequently use the identifier with the `removeAttachmentAsync` method to remove the attachment in the same session.
         *
         * **Important**: In recent builds of Outlook on Windows, a bug was introduced that incorrectly appends an `Authorization: Bearer` header to
         * this action (whether using this API or the Outlook UI). To work around this issue, you can try using the `addFileAttachmentFromBase64` API
         * introduced with requirement set 1.8.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * **Errors**:
         *
         * - `AttachmentSizeExceeded`: The attachment is larger than allowed.
         *
         * - `FileTypeNotSupported`: The attachment has an extension that is not allowed.
         *
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         *
         * @param uri - The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `isInline`: If true, indicates that the attachment will be shown inline in the message body,
         *                    and should not be displayed in the attachment list.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`.
         *                 On success, the attachment identifier will be provided in the `asyncResult.value` property.
         *                 If uploading the attachment fails, the `asyncResult` object will contain
         *                 an `Error` object that provides a description of the error.
         */
        addFileAttachmentAsync(uri: string, attachmentName: string, options: CommonAPI.AsyncContextOptions & { isInline: boolean }, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * You can subsequently use the identifier with the `removeAttachmentAsync` method to remove the attachment in the same session.
         *
         * **Important**: In recent builds of Outlook on Windows, a bug was introduced that incorrectly appends an `Authorization: Bearer` header to
         * this action (whether using this API or the Outlook UI). To work around this issue, you can try using the `addFileAttachmentFromBase64` API
         * introduced with requirement set 1.8.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * **Errors**:
         *
         * - `AttachmentSizeExceeded`: The attachment is larger than allowed.
         *
         * - `FileTypeNotSupported`: The attachment has an extension that is not allowed.
         *
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         *
         * @param uri - The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`.
         *                 On success, the attachment identifier will be provided in the `asyncResult.value` property.
         *                 If uploading the attachment fails, the `asyncResult` object will contain
         *                 an `Error` object that provides a description of the error.
         */
        addFileAttachmentAsync(uri: string, attachmentName: string, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The `addFileAttachmentFromBase64Async` method uploads the file from the Base64 encoding and attaches it to the item in the compose form.
         * This method returns the attachment identifier in the `asyncResult.value` object.
         *
         * You can subsequently use the identifier with the `removeAttachmentAsync` method to remove the attachment in the same session.
         *
         * **Note**: If you're using a data URL API (e.g., `readAsDataURL`), you need to strip out the data URL prefix then send the rest of the string to this API.
         * For example, if the full string is represented by `data:image/svg+xml;base64,<rest of Base64 string>`, remove `data:image/svg+xml;base64,`.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * **Errors**:
         *
         * - `AttachmentSizeExceeded`: The attachment is larger than allowed.
         *
         * - `FileTypeNotSupported`: The attachment has an extension that isn't allowed.
         *
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         *
         * **Note**: If you're adding an inline Base64 image to the body of a message or appointment being composed, you must first get the current item body using the
         * {@link https://learn.microsoft.com/javascript/api/outlook/office.body#outlook-office-body-getasync-member(1) | Office.context.mailbox.item.body.getAsync} 
         * method before inserting the image using `addFileAttachmentFromBase64Async`. Otherwise, the image won't render in the body once it's inserted.
         * For further guidance, see {@link https://learn.microsoft.com/office/dev/add-ins/outlook/add-and-remove-attachments-to-an-item-in-a-compose-form#attach-a-file | Attach a file}.
         *
         * @param base64File - The Base64-encoded content of an image or file to be added to an email or event. The maximum length of the encoded string is 27,892,122 characters (about 25 MB).
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `isInline`: If true, indicates that the attachment will be shown inline in the message body
         *                    and should not be displayed in the attachment list.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`.
         *                  On success, the attachment identifier will be provided in the `asyncResult.value` property.
         *                  If uploading the attachment fails, the `asyncResult` object will contain
         *                  an `Error` object that provides a description of the error.
         */
        addFileAttachmentFromBase64Async(base64File: string, attachmentName: string, options: CommonAPI.AsyncContextOptions &  { isInline: boolean }, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The `addFileAttachmentFromBase64Async` method uploads the file from the Base64 encoding and attaches it to the item in the compose form.
         * This method returns the attachment identifier in the `asyncResult.value` object.
         *
         * You can subsequently use the identifier with the `removeAttachmentAsync` method to remove the attachment in the same session.
         *
         * **Note**: If you're using a data URL API (e.g., `readAsDataURL`), you need to strip out the data URL prefix then send the rest of the string to this API.
         * For example, if the full string is represented by `data:image/svg+xml;base64,<rest of Base64 string>`, remove `data:image/svg+xml;base64,`.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * **Errors**:
         *
         * - `AttachmentSizeExceeded`: The attachment is larger than allowed.
         *
         * - `FileTypeNotSupported`: The attachment has an extension that isn't allowed.
         *
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         *
         * **Note**: If you're adding an inline Base64 image to the body of a message or appointment being composed, you must first get the current item body using the
         * {@link https://learn.microsoft.com/javascript/api/outlook/office.body#outlook-office-body-getasync-member(1) | Office.context.mailbox.item.body.getAsync} 
         * method before inserting the image using `addFileAttachmentFromBase64Async`. Otherwise, the image won't render in the body once it's inserted.
         * For further guidance, see {@link https://learn.microsoft.com/office/dev/add-ins/outlook/add-and-remove-attachments-to-an-item-in-a-compose-form#attach-a-file | Attach a file}.
         * 
         * @param base64File - The Base64-encoded content of an image or file to be added to an email or event. The maximum length of the encoded string is 27,892,122 characters (about 25 MB).
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`.
         *                  On success, the attachment identifier will be provided in the `asyncResult.value` property.
         *                  If uploading the attachment fails, the `asyncResult` object will contain
         *                  an `Error` object that provides a description of the error.
         */
        addFileAttachmentFromBase64Async(base64File: string, attachmentName: string, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Adds an event handler for a supported event. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Item object model
         * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param eventType - The event that should invoke the handler.
         * @param handler - The function to handle the event. The function must accept a single parameter, which is an object literal.
         *                The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        addHandlerAsync(eventType: CommonAPI.EventType | string, handler: any, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Adds an event handler for a supported event. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Item object model
         * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param eventType - The event that should invoke the handler.
         * @param handler - The function to handle the event. The function must accept a single parameter, which is an object literal.
         *                The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        addHandlerAsync(eventType: CommonAPI.EventType | string, handler: any, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message or appointment.
         *
         * The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form.
         * If you specify a callback function, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or
         * a code that indicates any error that occurred while attaching the item.
         * You can use the `options` parameter to pass state information to the callback function, if needed.
         *
         * You can subsequently use the identifier with the `removeAttachmentAsync` method to remove the attachment in the same session.
         *
         * If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that
         * you are editing; however, this is not supported and is not recommended.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * **Errors**:
         *
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         *
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of
         *                             type `Office.AsyncResult`.
         *                 On success, the attachment identifier will be provided in the `asyncResult.value` property.
         *                 If adding the attachment fails, the `asyncResult` object will contain
         *                 an `Error` object that provides a description of the error.
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message or appointment.
         *
         * The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form.
         * If you specify a callback function, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or
         * a code that indicates any error that occurred while attaching the item.
         * You can use the `options` parameter to pass state information to the callback function, if needed.
         *
         * You can subsequently use the identifier with the `removeAttachmentAsync` method to remove the attachment in the same session.
         *
         * If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that
         * you are editing; however, this is not supported and is not recommended.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * **Errors**:
         *
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         *
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter of
         *                             type `Office.AsyncResult`.
         *                 On success, the attachment identifier will be provided in the `asyncResult.value` property.
         *                 If adding the attachment fails, the `asyncResult` object will contain
         *                 an `Error` object that provides a description of the error.
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Closes the current item that is being composed
         *
         * The behaviors of the `close` method depends on the current state of the item being composed.
         * If the item has unsaved changes, the client prompts the user to save, discard, or close the action.
         *
         * In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.
         *
         * **Note**: In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save,
         * discard, or cancel even if no changes have occurred since the item was last saved.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         */
        close(): void;
        /**
         * Disables the Outlook client signature.
         *
         * For Windows and Mac rich clients, this API sets the signature under the "New Message" and "Replies/Forwards" sections
         * for the sending account to "(none)", effectively disabling the signature.
         * For Outlook on the web, the API should disable the signature option for new mails, replies, and forwards.
         * If the signature is selected, this API call should disable it.
         *
         * @remarks
         * [Api set: Mailbox 1.10]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        disableClientSignatureAsync(options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Disables the Outlook client signature.
         *
         * For Windows and Mac rich clients, this API sets the signature under the "New Message" and "Replies/Forwards" sections
         * for the sending account to "(none)", effectively disabling the signature.
         * For Outlook on the web, the API should disable the signature option for new mails, replies, and forwards.
         * If the signature is selected, this API call should disable it.
         *
         * @remarks
         * [Api set: Mailbox 1.10]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        disableClientSignatureAsync(callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Gets an attachment from a message or appointment and returns it as an `AttachmentContent` object.
         *
         * The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item. As a best practice, you should get
         * the attachment's identifier from a `getAttachmentsAsync` call, then in the same session, use that identifier to retrieve the attachment.
         * In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.
         * A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to
         * continue in a separate window.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * **Errors**:
         *
         * - `AttachmentTypeNotSupported`: The attachment type isn't supported. Unsupported types include embedded images in Rich Text Format,
         *                               or item attachment types other than email or calendar items (such as a contact or task item).
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param attachmentId - The identifier of the attachment you want to get.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object. If the call fails, the `asyncResult.error` property will contain
         *                an error code with the reason for the failure.
         */
        getAttachmentContentAsync(attachmentId: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<AttachmentContent>) => void): void;
        /**
         * Gets an attachment from a message or appointment and returns it as an `AttachmentContent` object.
         *
         * The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item. As a best practice, you should get
         * the attachment's identifier from a `getAttachmentsAsync` call, then in the same session, use that identifier to retrieve the attachment.
         * In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.
         * A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to
         * continue in a separate window.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * **Errors**:
         *
         * - `AttachmentTypeNotSupported`: The attachment type isn't supported. Unsupported types include embedded images in Rich Text Format,
         *                               or item attachment types other than email or calendar items (such as a contact or task item).
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param attachmentId - The identifier of the attachment you want to get.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object. If the call fails, the `asyncResult.error` property will contain
         *                an error code with the reason for the failure.
         */
        getAttachmentContentAsync(attachmentId: string, callback?: (asyncResult: CommonAPI.AsyncResult<AttachmentContent>) => void): void;
        /**
         * Gets the item's attachments as an array.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If the call fails, the `asyncResult.error` property will contain an error code with the reason for
         *                 the failure.
         */
        getAttachmentsAsync(options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<AttachmentDetailsCompose[]>) => void): void;
        /**
         * Gets the item's attachments as an array.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If the call fails, the `asyncResult.error` property will contain an error code with the reason for
         *                 the failure.
         */
        getAttachmentsAsync(callback?: (asyncResult: CommonAPI.AsyncResult<AttachmentDetailsCompose[]>) => void): void;
        /**
         * Gets initialization data passed when the add-in is {@link https://learn.microsoft.com/outlook/actionable-messages/invoke-add-in | activated by an actionable message}.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         *                 On success, the initialization context data is provided as a string (or an empty string if there's no initialization context)
         *                 in the `asyncResult.value` property.
         */
        getInitializationContextAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Gets initialization data passed when the add-in is {@link https://learn.microsoft.com/outlook/actionable-messages/invoke-add-in | activated by an actionable message}.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         *                 On success, the initialization context data is provided as a string (or an empty string if there's no initialization context)
         *                 in the `asyncResult.value` property.
         */
        getInitializationContextAsync(callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Asynchronously gets the ID of a saved item.
         *
         * When invoked, this method returns the item ID via the callback function.
         *
         * **Note**: If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API),
         * be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.
         * Until the item is synced, the `itemId` is not recognized and using it returns an error.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * **Errors**:
         *
         * - `ItemNotSaved`: The ID can't be retrieved until the item is saved.
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                   of type `Office.AsyncResult`.
         */
        getItemIdAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Asynchronously gets the ID of a saved item.
         *
         * When invoked, this method returns the item ID via the callback function.
         *
         * **Note**: If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API),
         * be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.
         * Until the item is synced, the `itemId` is not recognized and using it returns an error.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * **Errors**:
         *
         * - `ItemNotSaved`: The ID can't be retrieved until the item is saved.
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                   of type `Office.AsyncResult`.
         */
        getItemIdAsync(callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Asynchronously returns selected data from the subject or body of a message.
         *
         * If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data.
         * If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.
         *
         * To access the selected data from the callback function, call `asyncResult.value.data`.
         * To access the `source` property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.
         *
         * @returns
         * The selected data as a string with format determined by `coercionType`.
         *
         * @remarks
         * [Api set: Mailbox 1.2]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param coercionType - Requests a format for the data. If `Text`, the method returns the plain text as a string, removing any HTML tags present.
         *                     If `HTML`, the method returns the selected text, whether it is plaintext or HTML.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                   of type `Office.AsyncResult`.
         */
        getSelectedDataAsync(coercionType: CommonAPI.CoercionType | string, options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<any>) => void): void;
        /**
         * Asynchronously returns selected data from the subject or body of a message.
         *
         * If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data.
         * If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.
         *
         * To access the selected data from the callback function, call `asyncResult.value.data`.
         * To access the `source` property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.
         *
         * @returns
         * The selected data as a string with format determined by `coercionType`.
         *
         * @remarks
         * [Api set: Mailbox 1.2]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param coercionType - Requests a format for the data. If `Text`, the method returns the plain text as a string, removing any HTML tags present.
         *                     If `HTML`, the method returns the selected text, whether it is plaintext or HTML.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                   of type `Office.AsyncResult`.
         */
        getSelectedDataAsync(coercionType: CommonAPI.CoercionType | string, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Gets the properties of an appointment or message in a shared folder or shared mailbox (now in preview).
         *
         * For more information around using this API, see the
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/delegate-access | shared folders and shared mailbox} article.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         *                 The `value` property of the result is the properties of the shared item.
         */
        getSharedPropertiesAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<SharedProperties>) => void): void;
        /**
         * Gets the properties of an appointment or message in a shared folder or shared mailbox (now in preview).
         *
         * For more information around using this API, see the
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/delegate-access | shared folders and shared mailbox} article.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         *                 The `value` property of the result is the properties of the shared item.
         */
        getSharedPropertiesAsync(callback: (asyncResult: CommonAPI.AsyncResult<SharedProperties>) => void): void;
        /**
         * Gets if the client signature is enabled.
         *
         * For Windows and Mac rich clients, the API call should return `true` if the default signature for new messages, replies, or forwards is set
         * to a template for the sending Outlook account.
         * For Outlook on the web, the API call should return `true` if the signature is enabled for compose types `newMail`, `reply`, or `forward`.
         * If the settings are set to "(none)" in Mac or Windows rich clients or disabled in Outlook on the Web, the API call should return `false`.
         *
         * @remarks
         * [Api set: Mailbox 1.10]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                   type `Office.AsyncResult`.
         */
        isClientSignatureEnabledAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<boolean>) => void): void;
        /**
         * Gets if the client signature is enabled.
         *
         * For Windows and Mac rich clients, the API call should return `true` if the default signature for new messages, replies, or forwards is set
         * to a template for the sending Outlook account.
         * For Outlook on the web, the API call should return `true` if the signature is enabled for compose types `newMail`, `reply`, or `forward`.
         * If the settings are set to "(none)" in Mac or Windows rich clients or disabled in Outlook on the Web, the API call should return `false`.
         *
         * @remarks
         * [Api set: Mailbox 1.10]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                   type `Office.AsyncResult`.
         */
        isClientSignatureEnabledAsync(callback: (asyncResult: CommonAPI.AsyncResult<boolean>) => void): void;
        /**
         * Asynchronously loads custom properties for this add-in on the selected item.
         *
         * Custom properties are stored as key-value pairs on a per-app, per-item basis.
         * This method returns a {@link Office.CustomProperties | CustomProperties} object in the callback, which provides methods to access the custom properties specific to the
         * current item and the current add-in. Custom properties aren't encrypted on the item, so this shouldn't be used as secure storage.
         *
         * The custom properties are provided as a `CustomProperties` object in the `asyncResult.value` property.
         * This object can be used to get, set, save, and remove custom properties from the mail item.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * To learn more about custom properties, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/metadata-for-an-outlook-add-in | Get and set add-in metadata for an Outlook add-in}.
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         * @param userContext - Optional. Developers can provide any object they wish to access in the callback function.
         *                    This object can be accessed by the `asyncResult.asyncContext` property in the callback function.
         */
        loadCustomPropertiesAsync(callback: (asyncResult: CommonAPI.AsyncResult<CustomProperties>) => void, userContext?: any): void;
        /**
         * Removes an attachment from a message or appointment.
         *
         * The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.
         * As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment
         * in the same session. In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.
         * A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to
         * continue in a separate window.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * **Errors**:
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param attachmentId - The identifier of the attachment to remove. The maximum string length of the `attachmentId`
         *                       is 200 characters in Outlook on the web and on Windows.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                             type `Office.AsyncResult`.
         *                 If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.
         */
        removeAttachmentAsync(attachmentId: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Removes an attachment from a message or appointment.
         *
         * The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.
         * As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment
         * in the same session. In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.
         * A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to
         * continue in a separate window.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * **Errors**:
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param attachmentId - The identifier of the attachment to remove. The maximum string length of the `attachmentId`
         *                       is 200 characters in Outlook on the web and on Windows.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                             type `Office.AsyncResult`.
         *                 If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.
         */
        removeAttachmentAsync(attachmentId: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Removes the event handlers for a supported event type. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Item object model
         * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param eventType - The event that should revoke the handler.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        removeHandlerAsync(eventType: CommonAPI.EventType | string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Removes the event handlers for a supported event type. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Item object model
         * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * @param eventType - The event that should revoke the handler.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        removeHandlerAsync(eventType: CommonAPI.EventType | string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item ID via the callback function.
         * In Outlook on the web or Outlook in online mode, the item is saved to the server.
         * In Outlook in cached mode, the item is saved to the local cache.
         *
         * Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal
         * appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent.
         * Saving an existing appointment will send an update to added or removed attendees.
         *
         * When working with HTML-formatted content, it's important to note that the Outlook client may modify the content. This means that
         * subsequent calls to methods like `Body.getAsync`, `Body.setAsync`, and even `saveAsync` may not result in the same content.
         *
         * **Note**: If your add-in calls `saveAsync` on an item in compose mode in order to get an item ID to use with EWS or the REST API, be aware
         * that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.
         * Until the item is synced, using the item ID will return an error.
         *
         * **Note**: In Outlook on Mac, only build 16.35.308 or later supports saving a meeting.
         * Otherwise, the `saveAsync` method fails when called from a meeting in compose mode.
         * For a workaround, see {@link https://learn.microsoft.com/outlook/troubleshoot/calendars/cannot-save-meeting-as-draft-in-outlook-for-mac | Cannot save a meeting as a draft in Outlook for Mac by using Office JS API}.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * **Errors**:
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                   type `Office.AsyncResult`.
         */
        saveAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item ID via the callback function.
         * In Outlook on the web or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.
         *
         * Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal
         * appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent.
         * Saving an existing appointment will send an update to added or removed attendees.
         *
         * When working with HTML-formatted content, it's important to note that the Outlook client may modify the content. This means that
         * subsequent calls to methods like `Body.getAsync`, `Body.setAsync`, and even `saveAsync` may not result in the same content.
         *
         * **Note**: If your add-in calls `saveAsync` on an item in compose mode in order to get an item ID to use with EWS or the REST API, be aware that
         * when Outlook is in cached mode, it may take some time before the item is actually synced to the server.
         * Until the item is synced, using the item ID will return an error.
         *
         * **Note**: In Outlook on Mac, only build 16.35.308 or later supports saving a meeting.
         * Otherwise, the `saveAsync` method fails when called from a meeting in compose mode.
         * For a workaround, see {@link https://learn.microsoft.com/outlook/troubleshoot/calendars/cannot-save-meeting-as-draft-in-outlook-for-mac | Cannot save a meeting as a draft in Outlook for Mac by using Office JS API}.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * **Errors**:
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of
         *                   type `Office.AsyncResult`.
         */
        saveAsync(callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         *
         * The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is
         * selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned.
         * After insertion, the cursor is placed at the end of the inserted content.
         *
         * @remarks
         * [Api set: Mailbox 1.2]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * **Errors**:
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters.
         *             If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `coercionType`: If text, the current style is applied in Outlook on the web and Windows.
         *                      If the field is an HTML editor, only the text data is inserted, even if the data is HTML.
         *                      If html and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the
         *                      default style is applied in Outlook on desktop clients.
         *                      If the field is a text field, an `InvalidDataFormat` error is returned.
         *                      If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used;
         *                      if the field is text, then plain text is used.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        setSelectedDataAsync(data: string, options: CommonAPI.AsyncContextOptions & CoercionTypeOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         *
         * The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is
         * selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned.
         * After insertion, the cursor is placed at the end of the inserted content.
         *
         * @remarks
         * [Api set: Mailbox 1.2]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer
         *
         * **Errors**:
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters.
         *             If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        setSelectedDataAsync(data: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * The `AppointmentForm` object is used to access the currently selected appointment.
     *
     * @remarks
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface AppointmentForm {
        /**
         * Gets an object that provides methods for manipulating the body of an item.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        body: Body | string;
        /**
         * Gets or sets the date and time that the appointment is to end.
         *
         * The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the `convertToLocalClientTime` method to
         * convert the `end` property value to the client's local date and time.
         *
         * *Read mode*
         *
         * The `end` property returns a `Date` object.
         *
         * *Compose mode*
         *
         * The `end` property returns a `Time` object.
         *
         * When you use the `Time.setAsync` method to set the end time, you should use the `convertToUtcClientTime` method to convert the local time on
         * the client to UTC for the server.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        end: Time | Date;
        /**
        * Gets or sets the location of an appointment.
        *
        * *Read mode*
        *
        * The `location` property returns a string that contains the location of the appointment.
        *
        * *Compose mode*
        *
        * The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.
        *
        * @remarks
        *
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
        *
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
        */
       location: Location | string;
       /**
        * Provides access to the optional attendees of an event. The type of object and level of access depend on the mode of the current item.
        *
        * *Read mode*
        *
        * The `optionalAttendees` property returns an array that contains an {@link Office.EmailAddressDetails | EmailAddressDetails} object for
        * each optional attendee to the meeting. Collection size limits:
        *
        * - Windows: 500 members
        *
        * - Classic Mac UI: 100 members
        *
        * - New Mac UI, web browser, Android: No limit
        *
        * *Compose mode*
        *
        * The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the
        * optional attendees for a meeting. However, depending on the client/platform (i.e., Windows, Mac, etc.), limits may apply on how many
        * recipients you can get or update. See the {@link Office.Recipients | Recipients} object for more details.
        *
        * @remarks
        *
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
        *
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
        */
       optionalAttendees: Recipients[] | EmailAddressDetails[];
       /**
        * Provides access to the resources of an event. Returns an array of strings containing the resources required for the appointment.
        *
        * @remarks
        *
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
        *
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
        */
       resources: string[];
       /**
        * Provides access to the required attendees of an event. The type of object and level of access depend on the mode of the current item.
        *
        * *Read mode*
        *
        * The `requiredAttendees` property returns an array that contains an {@link Office.EmailAddressDetails | EmailAddressDetails} object for
        * each required attendee to the meeting. Collection size limits:
        *
        * - Windows: 500 members
        *
        * - Classic Mac UI: 100 members
        *
        * - New Mac UI, web browser, Android: No limit
        *
        * *Compose mode*
        *
        * The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the
        * required attendees for a meeting. However, depending on the client/platform (i.e., Windows, Mac, etc.), limits may apply on how many
        * recipients you can get or update. See the {@link Office.Recipients | Recipients} object for more details.
        *
        * @remarks
        *
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
        *
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
        */
       requiredAttendees: Recipients[] | EmailAddressDetails[];
       /**
        * Gets or sets the date and time that the appointment is to begin.
        *
        * The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the `convertToLocalClientTime` method
        * to convert the value to the client's local date and time.
        *
        * *Read mode*
        *
        * The `start` property returns a `Date` object.
        *
        * *Compose mode*
        *
        * The `start` property returns a `Time` object.
        *
        * When you use the `Time.setAsync` method to set the start time, you should use the `convertToUtcClientTime` method to convert the local time on
        * the client to UTC for the server.
        *
        * @remarks
        *
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
        *
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
        */
       start: Time | Date;
       /**
        * Gets or sets the description that appears in the subject field of an item.
        *
        * The `subject` property gets or sets the entire subject of the item, as sent by the email server.
        *
        * *Read mode*
        *
        * The `subject` property returns a string. Use the `normalizedSubject` property to get the subject minus any leading prefixes such as RE: and FW:.
        *
        * *Compose mode*
        *
        * The `subject` property returns a `Subject` object that provides methods to get and set the subject.
        *
        * @remarks
        *
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
        *
        * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
        */
       subject: Subject | string;
    }
    /**
     * The appointment attendee mode of {@link Office.Item | Office.context.mailbox.item}.
     *
     * **Important**: This is an internal Outlook object, not directly exposed through existing interfaces.
     * You should treat this as a mode of `Office.context.mailbox.item`. For more information, refer to the
     * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item | Object Model} page.
     *
     * Parent interfaces:
     *
     * - {@link Office.ItemRead | ItemRead}
     *
     * - {@link Office.Appointment | Appointment}
     */
    export interface AppointmentRead extends Appointment, ItemRead {
        /**
         * Gets the item's attachments as an array.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * **Note**: Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned. For more information, see
         * {@link https://support.microsoft.com/office/434752e1-02d3-4e90-9124-8b81e49a8519 | Blocked attachments in Outlook}.
         *
         */
        attachments: AttachmentDetails[];
        /**
         * Gets an object that provides methods for manipulating the body of an item.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        body: Body;
        /**
         * Gets an object that provides methods for managing the item's categories.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        categories: Categories;
        /**
         * Gets the date and time that an item was created.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        dateTimeCreated: Date;
        /**
         * Gets the date and time that an item was last modified.
         *
         * **Note**: This member is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        dateTimeModified: Date;
        /**
         * Gets the date and time that the appointment is to end.
         *
         * The `end` property is a `Date` object expressed as a Coordinated Universal Time (UTC) date and time value.
         * You can use the `convertToLocalClientTime` method to convert the `end` property value to the client's local date and time.
         *
         * When you use the `Time.setAsync` method to set the end time, you should use the `convertToUtcClientTime` method to convert the local time on
         * the client to UTC for the server.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        end: Date;
        /**
         * Gets the locations of an appointment.
         *
         * The `enhancedLocation` property returns an {@link Office.EnhancedLocation | EnhancedLocation} object that allows you to get the set of locations
         * (each represented by a {@link Office.LocationDetails | LocationDetails} object) associated with the appointment.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        enhancedLocation: EnhancedLocation;
        /**
         * Returns a boolean value indicating whether the event is all day.
         *
         * @remarks
         * [Api set: Mailbox preview]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**:  **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**:  Appointment Attendee
         *
         * @beta
         */
        isAllDayEvent: boolean;
        /**
         * Gets the Exchange Web Services item class of the selected item.
         *
         * You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.
         *
         * <table>
         *   <tr>
         *     <th>Type</th>
         *     <th>Description</th>
         *     <th>Item Class</th>
         *   </tr>
         *   <tr>
         *     <td>Appointment items</td>
         *     <td>These are calendar items of the item class IPM.Appointment or IPM.Appointment.Occurrence.</td>
         *     <td>IPM.Appointment, IPM.Appointment.Occurrence</td>
         *   </tr>
         *   <tr>
         *     <td>Message items</td>
         *     <td>These include email messages that have the default message class IPM.Note, and meeting requests, responses, and cancellations, that use IPM.Schedule.Meeting as the base message class.</td>
         *     <td>IPM.Note, IPM.Schedule.Meeting.Request, IPM.Schedule.Meeting.Neg, IPM.Schedule.Meeting.Pos, IPM.Schedule.Meeting.Tent, IPM.Schedule.Meeting.Canceled</td>
         *   </tr>
         * </table>
         *
         */
        itemClass: string;
        /**
         * Gets the {@link https://learn.microsoft.com/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange | Exchange Web Services item identifier}
         * for the current item.
         *
         * The `itemId` property is not available in compose mode.
         * If an item identifier is required, the `saveAsync` method can be used to save the item to the store, which will return the item identifier
         * in the `asyncResult.value` parameter in the callback function.
         *
         * **Note**: The identifier returned by the `itemId` property is the same as the
         * {@link https://learn.microsoft.com/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange | Exchange Web Services item identifier}.
         * The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.
         * Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`.
         * For more details, see {@link https://learn.microsoft.com/office/dev/add-ins/outlook/use-rest-api#get-the-item-id | Use the Outlook REST APIs from an Outlook add-in}.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        itemId: string;
        /**
         * Gets the type of item that an instance represents.
         *
         * The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the item object instance is a message or an appointment.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        itemType: MailboxEnums.ItemType | string;
        /**
         * Gets the location of an appointment.
         *
         * The `location` property returns a string that contains the location of the appointment.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        location: string;
        /**
         * Gets the subject of an item, with all prefixes removed (including RE: and FWD:).
         *
         * The `normalizedSubject` property gets the subject of the item, with any standard prefixes (such as RE: and FW:) that are added by email programs.
         * To get the subject of the item with the prefixes intact, use the `subject` property.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        normalizedSubject: string;
        /**
         * Gets the notification messages for an item.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        notificationMessages: NotificationMessages;
        /**
         * Provides access to the optional attendees of an event. The type of object and level of access depend on the mode of the current item.
         *
         * The `optionalAttendees` property returns an array that contains an {@link Office.EmailAddressDetails | EmailAddressDetails} object for
         * each optional attendee to the meeting. Collection size limits:
         *
         * - Windows: 500 members
         *
         * - Classic Mac UI: 100 members
         *
         * - New Mac UI, web browser, Android: No limit
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        optionalAttendees: EmailAddressDetails[];
        /**
         * Gets the meeting organizer's email properties.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        organizer: EmailAddressDetails;
        /**
         * Gets the recurrence pattern of an appointment. Gets the recurrence pattern of a meeting request.
         *
         * The `recurrence` property returns a {@link Office.Recurrence | Recurrence} object for recurring appointments or meetings requests
         * if an item is a series or an instance in a series. `null` is returned for single appointments and meeting requests of single appointments.
         *
         * **Note**: Meeting requests have an `itemClass` value of `IPM.Schedule.Meeting.Request`.
         *
         * **Note**: If the recurrence object is null, this indicates that the object is a single appointment or a meeting request of a single
         * appointment and NOT a part of a series.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        recurrence: Recurrence;
        /**
         * Provides access to the required attendees of an event. The type of object and level of access depend on the mode of the current item.
         *
         * The `requiredAttendees` property returns an array that contains an {@link Office.EmailAddressDetails | EmailAddressDetails} object for
         * each required attendee to the meeting. Collection size limits:
         *
         * - Windows: 500 members
         *
         * - Classic Mac UI: 100 members
         *
         * - New Mac UI, web browser, Android: No limit
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        requiredAttendees: EmailAddressDetails[];
        /**
         * Gets the date and time that the appointment is to begin.
         *
         * The `start` property is a `Date` object expressed as a Coordinated Universal Time (UTC) date and time value.
         * You can use the `convertToLocalClientTime` method to convert the value to the client's local date and time.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        start: Date;
        /**
         * Gets the ID of the series that an instance belongs to.
         *
         * In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item
         * that this item belongs to. However, on iOS and Android, the seriesId returns the REST ID of the parent item.
         *
         * **Note**: The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.
         * The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API. Before making REST API calls using this value, it
         * should be converted using `Office.context.mailbox.convertToRestId`.
         * For more details, see {@link https://learn.microsoft.com/office/dev/add-ins/outlook/use-rest-api | Use the Outlook REST APIs from an Outlook add-in}.
         *
         * The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests
         * and returns `undefined` for any other items that are not meeting requests.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        seriesId: string;
        /**
         * Gets the description that appears in the subject field of an item.
         *
         * The `subject` property gets or sets the entire subject of the item, as sent by the email server.
         *
         * The `subject` property returns a string. Use the `normalizedSubject` property to get the subject minus any leading prefixes such as RE: and FW:.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        subject: string;
        /**
         * Provides the sensitivity value of the appointment.
         *
         * @remarks
         * [Api set: Mailbox preview]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**:  **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**:  Appointment Attendee
         *
         * @beta
         */
        sensitivity: MailboxEnums.AppointmentSensitivityType;

        /**
         * Adds an event handler for a supported event. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Item object model
         * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param eventType - The event that should invoke the handler.
         * @param handler - The function to handle the event. The function must accept a single parameter, which is an object literal.
         *                The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        addHandlerAsync(eventType: CommonAPI.EventType | string, handler: any, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Adds an event handler for a supported event. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Item object model
         * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param eventType - The event that should invoke the handler.
         * @param handler - The function to handle the event. The function must accept a single parameter, which is an object literal.
         *                The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        addHandlerAsync(eventType: CommonAPI.EventType | string, handler: any, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Displays a reply form that includes either the sender and all recipients of the selected message or the organizer and all attendees of the
         * selected appointment.
         *
         * In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2-column or 1-column view.
         *
         * If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.
         *
         * When attachments are specified in the `formData.attachments` parameter, Outlook attempts to download all attachments and attach them to the
         * reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *                   OR a {@link Office.ReplyFormData | ReplyFormData} object that contains body or attachment data and a callback function.
         */
        displayReplyAllForm(formData: string | ReplyFormData): void;
        /**
         * Displays a reply form that includes either the sender and all recipients of the selected message or the organizer and all attendees of the
         * selected appointment.
         *
         * In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2-column or 1-column view.
         *
         * If any of the string parameters exceed their limits, `displayReplyAllFormAsync` throws an exception.
         *
         * When attachments are specified in the `formData.attachments` parameter, Outlook attempts to download all attachments and attach them to the
         * reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.9]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *                   OR a {@link Office.ReplyFormData | ReplyFormData} object that contains body or attachment data and a callback function.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        displayReplyAllFormAsync(formData: string | ReplyFormData, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Displays a reply form that includes either the sender and all recipients of the selected message or the organizer and all attendees of the
         * selected appointment.
         *
         * In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2-column or 1-column view.
         *
         * If any of the string parameters exceed their limits, `displayReplyAllFormAsync` throws an exception.
         *
         * When attachments are specified in the `formData.attachments` parameter, Outlook attempts to download all attachments and attach them to the
         * reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.9]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *                   OR a {@link Office.ReplyFormData | ReplyFormData} object that contains body or attachment data and a callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        displayReplyAllFormAsync(formData: string | ReplyFormData, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.
         *
         * In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2-column or 1-column view.
         *
         * If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.
         *
         * When attachments are specified in the `formData.attachments` parameter, Outlook attempts to download all attachments and attach them to the
         * reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *                   OR a {@link Office.ReplyFormData | ReplyFormData} object that contains body or attachment data and a callback function.
         */
        displayReplyForm(formData: string | ReplyFormData): void;
        /**
         * Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.
         *
         * In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2-column or 1-column view.
         *
         * If any of the string parameters exceed their limits, `displayReplyFormAsync` throws an exception.
         *
         * When attachments are specified in the `formData.attachments` parameter, Outlook attempts to download all attachments and attach them to the
         * reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.9]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *                   OR a {@link Office.ReplyFormData | ReplyFormData} object that contains body or attachment data and a callback function.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        displayReplyFormAsync(formData: string | ReplyFormData, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.
         *
         * In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2-column or 1-column view.
         *
         * If any of the string parameters exceed their limits, `displayReplyFormAsync` throws an exception.
         *
         * When attachments are specified in the `formData.attachments` parameter, Outlook attempts to download all attachments and attach them to the
         * reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.9]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *                   OR a {@link Office.ReplyFormData | ReplyFormData} object that contains body or attachment data and a callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        displayReplyFormAsync(formData: string | ReplyFormData, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Gets an attachment from a message or appointment and returns it as an `AttachmentContent` object.
         *
         * The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item. As a best practice, you should get
         * the attachment's identifier from an {@link Office.AppointmentRead.attachments | item.attachments} call, then in the same session, use that identifier
         * to retrieve the attachment. In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.
         * A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to
         * continue in a separate window.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * **Errors**:
         *
         * - `AttachmentTypeNotSupported`: The attachment type isn't supported. Unsupported types include embedded images in Rich Text Format,
         *                               or item attachment types other than email or calendar items (such as a contact or task item).
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param attachmentId - The identifier of the attachment you want to get.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object. If the call fails, the `asyncResult.error` property will contain
         *                 an error code with the reason for the failure.
         */
        getAttachmentContentAsync(attachmentId: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<AttachmentContent>) => void): void;
        /**
         * Gets an attachment from a message or appointment and returns it as an `AttachmentContent` object.
         *
         * The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item. As a best practice, you should get
         * the attachment's identifier from an {@link Office.AppointmentRead.attachments | item.attachments} call, then in the same session, use that identifier
         * to retrieve the attachment. In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.
         * A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to
         * continue in a separate window.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * **Errors**:
         *
         * - `AttachmentTypeNotSupported`: The attachment type isn't supported. Unsupported types include embedded images in Rich Text Format,
         *                               or item attachment types other than email or calendar items (such as a contact or task item).
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param attachmentId - The identifier of the attachment you want to get.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object. If the call fails, the `asyncResult.error` property will contain
         *                 an error code with the reason for the failure.
         */
        getAttachmentContentAsync(attachmentId: string, callback?: (asyncResult: CommonAPI.AsyncResult<AttachmentContent>) => void): void;
        /**
         * Gets the entities found in the selected item's body.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        getEntities(): Entities;
        /**
         * Gets an array of all the entities of the specified entity type found in the selected item's body.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @returns
         * If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.
         * If no entities of the specified type are present in the item's body, the method returns an empty array.
         * Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param entityType - One of the `EntityType` enumeration values.
         *
         * While the minimum permission level to use this method is **restricted**, some entity types require **read item** to access, as specified in the following table.
         *
         * <table>
         *   <tr>
         *     <th>Value of entityType</th>
         *     <th>Type of objects in returned array</th>
         *     <th>Required Permission Level</th>
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
        getEntitiesByType(entityType: MailboxEnums.EntityType | string): (string | Contact | MeetingSuggestion | PhoneNumber | TaskSuggestion)[];
        /**
         * Returns well-known entities in the selected item that pass the named filter defined in an XML manifest file.
         * 
         * **Note**: This method is used with the {@link https://learn.microsoft.com/office/dev/add-ins/outlook/activation-rules | activation rules feature for Outlook add-ins}, which isn't supported by the {@link https://learn.microsoft.com/office/dev/add-ins/develop/json-manifest-overview | Teams manifest for Office Add-ins (preview)}.
         *
         * The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the `ItemHasKnownEntity` rule element
         * in the manifest XML file with the specified `FilterName` element value.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @returns If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter,
         * the method returns `null`.
         * If the `name` parameter matches an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match,
         * the method return an empty array.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param name - The name of the `ItemHasKnownEntity` rule element that defines the filter to match.
         */
        getFilteredEntitiesByName(name: string): (string | Contact | MeetingSuggestion | PhoneNumber | TaskSuggestion)[];
        /**
         * Gets initialization data passed when the add-in is {@link https://learn.microsoft.com/outlook/actionable-messages/invoke-add-in | activated by an actionable message}.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         *                 On success, the initialization context data is provided as a string (or an empty string if there's no initialization context)
         *                 in the `asyncResult.value` property.
         */
        getInitializationContextAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Gets initialization data passed when the add-in is {@link https://learn.microsoft.com/outlook/actionable-messages/invoke-add-in | activated by an actionable message}.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         *                 On success, the initialization context data is provided as a string (or an empty string if there's no initialization context)
         *                 in the `asyncResult.value` property.
         */
        getInitializationContextAsync(callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Returns string values in the selected item that match the regular expressions defined in an XML manifest file.
         * 
         * **Note**: This method is used with the {@link https://learn.microsoft.com/office/dev/add-ins/outlook/activation-rules | activation rules feature for Outlook add-ins}, which isn't supported by the {@link https://learn.microsoft.com/office/dev/add-ins/develop/json-manifest-overview | Teams manifest for Office Add-ins (preview)}.
         *
         * The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or
         * `ItemHasKnownEntity` rule element in the manifest XML file.
         * For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule.
         * The `PropertyName` simple type defines the supported properties.
         *
         * If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body
         * and should not attempt to return the entire body of the item.
         * Using a regular expression such as .* to obtain the entire body of an item does not always return the expected results.
         * Instead, use the `Body.getAsync` method to retrieve the entire body.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @returns
         * An object that contains arrays of strings that match the regular expressions defined in the manifest XML file.
         * The name of each array is equal to the corresponding value of the RegExName attribute of the matching `ItemHasRegularExpressionMatch` rule
         * or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        getRegExMatches(): any;
        /**
         * Returns string values in the selected item that match the named regular expression defined in an XML manifest file.
         * 
         * **Note**: This method is used with the {@link https://learn.microsoft.com/office/dev/add-ins/outlook/activation-rules | activation rules feature for Outlook add-ins}, which isn't supported by the {@link https://learn.microsoft.com/office/dev/add-ins/develop/json-manifest-overview | Teams manifest for Office Add-ins (preview)}.
         *
         * The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule
         * element in the manifest XML file with the specified `RegExName` element value.
         *
         * If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body
         * and should not attempt to return the entire body of the item.
         * Using a regular expression such as .* to obtain the entire body of an item does not always return the expected results.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @returns
         * An array that contains the strings that match the regular expression defined in the manifest XML file.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param name - The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.
         */
        getRegExMatchesByName(name: string): string[];
        /**
         * Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to contextual add-ins.
         * 
         * **Note**: This method is used with the {@link https://learn.microsoft.com/office/dev/add-ins/outlook/activation-rules | activation rules feature for Outlook add-ins}, which isn't supported by the {@link https://learn.microsoft.com/office/dev/add-ins/develop/json-manifest-overview | Teams manifest for Office Add-ins (preview)}.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.6]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param name - The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.
         */
        getSelectedEntities(): Entities;
        /**
         * Returns string values in a highlighted match that match the regular expressions defined in an XML manifest file.
         * Highlighted matches apply to contextual add-ins.
         * 
         * **Note**: This method is used with the {@link https://learn.microsoft.com/office/dev/add-ins/outlook/activation-rules | activation rules feature for Outlook add-ins}, which isn't supported by the {@link https://learn.microsoft.com/office/dev/add-ins/develop/json-manifest-overview | Teams manifest for Office Add-ins (preview)}.
         *
         * The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or
         * `ItemHasKnownEntity` rule element in the manifest XML file.
         * For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule.
         * The `PropertyName` simple type defines the supported properties.
         *
         * If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body
         * and should not attempt to return the entire body of the item.
         * Using a regular expression such as .* to obtain the entire body of an item does not always return the expected results.
         * Instead, use the `Body.getAsync` method to retrieve the entire body.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @returns
         * An object that contains arrays of strings that match the regular expressions defined in the manifest XML file.
         * The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule
         * or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.
         *
         * @remarks
         * [Api set: Mailbox 1.6]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         */
        getSelectedRegExMatches(): any;
        /**
         * Gets the properties of an appointment or message in a shared folder or shared mailbox (now in preview).
         *
         * For more information around using this API, see the
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/delegate-access | shared folders and shared mailbox} article.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         *                 The `value` property of the result is the properties of the shared item.
         */
        getSharedPropertiesAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<SharedProperties>) => void): void;
        /**
         * Gets the properties of an appointment or message in a shared folder or shared mailbox (now in preview).
         *
         * For more information around using this API, see the
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/delegate-access | shared folders and shared mailbox} article.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         *                 The `value` property of the result is the properties of the shared item.
         */
        getSharedPropertiesAsync(callback: (asyncResult: CommonAPI.AsyncResult<SharedProperties>) => void): void;
        /**
         * Asynchronously loads custom properties for this add-in on the selected item.
         *
         * Custom properties are stored as key-value pairs on a per-app, per-item basis.
         * This method returns a {@link Office.CustomProperties | CustomProperties} object in the callback, which provides methods to access the custom properties specific to the
         * current item and the current add-in. Custom properties aren't encrypted on the item, so this shouldn't be used as secure storage.
         *
         * The custom properties are provided as a `CustomProperties` object in the `asyncResult.value` property.
         * This object can be used to get, set, save, and remove custom properties from the mail item.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * To learn more about custom properties, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/metadata-for-an-outlook-add-in | Get and set add-in metadata for an Outlook add-in}.
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         * @param userContext - Optional. Developers can provide any object they wish to access in the callback function.
         *                    This object can be accessed by the `asyncResult.asyncContext` property in the callback function.
         */
        loadCustomPropertiesAsync(callback: (asyncResult: CommonAPI.AsyncResult<CustomProperties>) => void, userContext?: any): void;
        /**
         * Removes the event handlers for a supported event type. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Item object model
         * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param eventType - The event that should revoke the handler.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        removeHandlerAsync(eventType: CommonAPI.EventType | string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Removes the event handlers for a supported event type. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Item object model
         * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Attendee
         *
         * @param eventType - The event that should revoke the handler.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        removeHandlerAsync(eventType: CommonAPI.EventType | string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * Provides the current dates and times of the appointment that raised the `Office.EventType.AppointmentTimeChanged` event.
     *
     * @remarks
     * [Api set: Mailbox 1.7]
     */
    export interface AppointmentTimeChangedEventArgs {
        /**
         * Gets the appointment end date and time.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         */
        end: Date;
        /**
         * Gets the appointment start date and time.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         */
        start: Date;
        /**
         * Gets the type of the event. For details, refer to {@link https://learn.microsoft.com/javascript/api/office/office.eventtype | Office.EventType}.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         */
        type: "olkAppointmentTimeChanged";
    }
    /**
     * Represents the content of an attachment on a message or appointment item.
     *
     * @remarks
     * [Api set: Mailbox 1.8]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface AttachmentContent {
        /**
         * The content of an attachment as a string.
         */
        content: string;
        /**
         * The string format to use for an attachment's content.
         *
         * For file attachments, the formatting is a base64-encoded string.
         *
         * For item attachments that represent messages and were attached by drag-and-drop or "Attach Item",
         * the formatting is a string representing an .eml formatted file.
         * **Important**: If a message item was attached by drag-and-drop in Outlook on the web, then `getAttachmentContentAsync` throws an error.
         *
         * For item attachments that represent calendar items and were attached by drag-and-drop or "Attach Item",
         * the formatting is a string representing an .icalendar file.
         * **Important**: If a calendar item was attached by drag-and-drop in Outlook on the web, then `getAttachmentContentAsync` throws an error.
         *
         * For cloud attachments, the formatting is a URL string.
         */
        format: MailboxEnums.AttachmentContentFormat | string;
    }
    /**
     * Represents an attachment on an item. Compose mode only.
     *
     * An array of `AttachmentDetailsCompose` objects is returned as the attachments property of an appointment or message item.
     *
     * @remarks
     * [Api set: Mailbox 1.8]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
     */
    export interface AttachmentDetailsCompose {
        /**
         * Gets a value that indicates the type of an attachment.
         */
        attachmentType: MailboxEnums.AttachmentType | string;
        /**
         * Gets the index of the attachment.
         */
        id: string;
        /**
         * Gets a value that indicates whether the attachment should be displayed in the body of the item.
         */
        isInline: boolean;
        /**
         * Gets the name of the attachment.
         *
         * **Important**: For message or appointment items that were attached by drag-and-drop or "Attach Item",
         * `name` includes a file extension in Outlook on Mac, but excludes the extension on the web or Windows.
         */
        name: string;
        /**
         * Gets the size of the attachment in bytes.
         */
        size: number;
        /**
         * Gets the url of the attachment if its type is `MailboxEnums.AttachmentType.Cloud`.
         */
        url?: string;
    }
    /**
     * Represents an attachment on an item from the server. Read mode only.
     *
     * An array of `AttachmentDetails` objects is returned as the attachments property of an appointment or message item.
     *
     * @remarks
     * [Api set: Mailbox 1.1]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
     */
    export interface AttachmentDetails {
        /**
         * Gets a value that indicates the type of an attachment.
         */
        attachmentType: MailboxEnums.AttachmentType | string;
        /**
         * Gets the MIME content type of the attachment.
         *
         * **Warning**: While the `contentType` value is a direct lookup of the attachment's extension, the internal mapping isn't actively maintained
         * so this property has been deprecated. If you require specific types, grab the attachment's extension and process accordingly. For details,
         * refer to the {@link https://devblogs.microsoft.com/microsoft365dev/outlook-javascript-api-deprecation-for-attachmentdetails-contenttype-property/ | related blog post }.
         *
         * @deprecated If you require specific content types, grab the attachment's extension and process accordingly.
         */
        contentType: string;
        /**
         * Gets the Exchange attachment ID of the attachment.
         * However, if the attachment type is `MailboxEnums.AttachmentType.Cloud`, then a URL for the file is returned.
         */
        id: string;
        /**
         * Gets a value that indicates whether the attachment should be displayed in the body of the item.
         */
        isInline: boolean;
        /**
         * Gets the name of the attachment.
         *
         * **Important**: For message or appointment items that were attached by drag-and-drop or "Attach Item",
         * `name` includes a file extension in Outlook on Mac, but excludes the extension on the web or Windows.
         */
        name: string;
        /**
         * Gets the size of the attachment in bytes.
         */
        size: number;
    }
    /**
     * Provides information about the attachments that raised the `Office.EventType.AttachmentsChanged` event.
     *
     * @remarks
     * [Api set: Mailbox 1.8]
     */
    export interface AttachmentsChangedEventArgs {
        /**
         * Represents the set of attachments that were added or removed.
         * For each such attachment, gets `id`, `name`, `size`, and `attachmentType` properties.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         */
        attachmentDetails: object[];
        /**
         * Gets whether the attachments were added or removed. For details, refer to {@link Office.MailboxEnums.AttachmentStatus | MailboxEnums.AttachmentStatus}.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         */
        attachmentStatus: MailboxEnums.AttachmentStatus | string;
        /**
         * Gets the type of the event. For details, refer to {@link https://learn.microsoft.com/javascript/api/office/office.eventtype | Office.EventType}.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         */
        type: "olkAttachmentsChanged";
    }
    /**
     * The body object provides methods for adding and updating the content of the message or appointment.
     * It is returned in the body property of the selected item.
     *
     * @remarks
     * [Api set: Mailbox 1.1]
     *
     * **Known issue with HTML table border colors**
     *
     * Outlook on Windows: If you're setting various cell borders to different colors in an HTML table in Compose mode, a cell's borders may not reflect
     * the expected color. For the known behavior, visit {@link https://github.com/OfficeDev/office-js/issues/1818 | OfficeDev/office-js issue #1818}.
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface Body {
        /**
         * Appends on send the specified content to the end of the item body, after any signature.
         *
         * To use `appendOnSendAsync`, you must specify a supplementary permission in the manifest. Details vary with the type of manifest. See {@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Understanding Outlook add-in permissions}. To learn more
         * about append-on-send and its configuration, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/append-on-send | Implement append-on-send in your Outlook add-in}.
         * 
         * **Important**: If the {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-on-send-addins?tabs=windows | on-send feature}
         * is implemented with append-on-send, the following apply.
         * 
         * - If the user is running add-ins that implement the on-send feature using `ItemSend` in the manifest, append-on-send runs before on-send functionality.
         *
         * - If your add-in implements the on-send feature and calls `appendOnSendAsync` in the `ItemSend` handler,
         * the `appendOnSendAsync` call returns an error as this scenario isn't supported.
         *
         * **Recommended**: Call `getTypeAsync` then pass the returned value to the `options.coercionType` parameter.
         *
         * **Note**: To clear data from a previous `appendOnSendAsync` call, you can call it again with the `data` parameter set to `null`.
         *
         * @remarks
         * [Api set: Mailbox 1.9]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `DataExceedsMaximumSize`: The `data` parameter is longer than 5,000 characters.
         *
         * - `InvalidFormatError`: The `options.coercionType` parameter is set to `Office.CoercionType.Html` but the message body is in plain text.
         *
         * @param data - The string to be added to the end of the body. The string is limited to 5,000 characters.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `coercionType`: The desired format for the data to be appended. The string in the `data` parameter will be converted to this format.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`. Any errors encountered will be provided in the `asyncResult.error` property.
         */
        appendOnSendAsync(data: string, options: CommonAPI.AsyncContextOptions & CoercionTypeOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Appends on send the specified content to the end of the item body, after any signature.
         *
         * To use `appendOnSendAsync`, you must specify a supplementary permission in the manifest. Details vary with the type of manifest. See {@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Understanding Outlook add-in permissions}. To learn more
         * about append-on-send and its configuration, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/append-on-send | Implement append-on-send in your Outlook add-in}.
         * 
         * **Important**: If the {@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-on-send-addins?tabs=windows | on-send feature}
         * is implemented with append-on-send, the following apply.
         * 
         * - If the user is running add-ins that implement the on-send feature using `ItemSend` in the manifest, append-on-send runs before on-send functionality.
         *
         * - If your add-in implements the on-send feature and calls `appendOnSendAsync` in the `ItemSend` handler,
         * the `appendOnSendAsync` call returns an error as this scenario isn't supported.
         *
         * **Recommended**: Call `getTypeAsync` then pass the returned value to the `options.coercionType` parameter.
         *
         * **Note**: To clear data from a previous `appendOnSendAsync` call, you can call it again with the `data` parameter set to `null`.
         *
         * @remarks
         * [Api set: Mailbox 1.9]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `DataExceedsMaximumSize`: The `data` parameter is longer than 5,000 characters.
         *
         * - `InvalidFormatError`: The `options.coercionType` parameter is set to `Office.CoercionType.Html` but the message body is in plain text.
         *
         * @param data - The string to be added to the end of the body. The string is limited to 5,000 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`. Any errors encountered will be provided in the `asyncResult.error` property.
         */
        appendOnSendAsync(data: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Returns the current body in a specified format.
         *
         * This method returns the entire current body in the format specified by `coercionType`.
         *
         * When working with HTML-formatted bodies, it is important to note that the `Body.getAsync` and `Body.setAsync` methods are not idempotent.
         * The value returned from the `getAsync` method will not necessarily be exactly the same as the value that was passed in the `setAsync` method previously.
         * The client may modify the value passed to `setAsync` in order to make it render efficiently with its rendering engine.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param coercionType - The format for the returned body.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type Office.AsyncResult. The body is provided in the requested format in the `asyncResult.value` property.
         */
        getAsync(coercionType: CommonAPI.CoercionType | string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Returns the current body in a specified format.
         *
         * This method returns the entire current body in the format specified by `coercionType`.
         *
         * When working with HTML-formatted bodies, it is important to note that the `Body.getAsync` and `Body.setAsync` methods are not idempotent.
         * The value returned from the `getAsync` method will not necessarily be exactly the same as the value that was passed in the `setAsync` method previously.
         * The client may modify the value passed to `setAsync` in order to make it render efficiently with its rendering engine.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param coercionType - The format for the returned body.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type Office.AsyncResult. The body is provided in the requested format in the `asyncResult.value` property.
         */
        getAsync(coercionType: CommonAPI.CoercionType | string, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Gets a value that indicates whether the content is in HTML or text format.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`.
         *                  The content type is returned as one of the `CoercionType` values in the `asyncResult.value` property.
         */
        getTypeAsync(options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<CommonAPI.CoercionType>) => void): void;
        /**
         * Gets a value that indicates whether the content is in HTML or text format.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`.
         *                  The content type is returned as one of the `CoercionType` values in the `asyncResult.value` property.
         */
        getTypeAsync(callback?: (asyncResult: CommonAPI.AsyncResult<CommonAPI.CoercionType>) => void): void;
        /**
         * Adds the specified content to the beginning of the item body.
         *
         * The `prependAsync` method inserts the specified string at the beginning of the item body.
         * After insertion, the cursor is returned to its original place, relative to the inserted content.
         *
         * When working with HTML-formatted bodies, it's important to note that the client may modify the value passed to `prependAsync` in order to
         * make it render efficiently with its rendering engine. This means that the value returned from a subsequent call to the `Body.getAsync` method
         * (introduced in Mailbox 1.3) will not necessarily exactly contain the value that was passed in the `prependAsync` method previously.
         *
         * When including links in HTML markup, you can disable online link preview by setting the `id` attribute on the anchor (\<a\>) to "LPNoLP"
         * (see the **Examples** section for a sample).
         *
         * **Recommended**: Call `getTypeAsync` then pass the returned value to the `options.coercionType` parameter.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `DataExceedsMaximumSize`: The data parameter is longer than 1,000,000 characters.
         *
         * @param data - The string to be inserted at the beginning of the body. The string is limited to 1,000,000 characters.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `coercionType`: The desired format for the body. The string in the `data` parameter will be converted to this format.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`. Any errors encountered will be provided in the `asyncResult.error` property.
         */
        prependAsync(data: string, options: CommonAPI.AsyncContextOptions & CoercionTypeOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Adds the specified content to the beginning of the item body.
         *
         * The `prependAsync` method inserts the specified string at the beginning of the item body.
         * After insertion, the cursor is returned to its original place, relative to the inserted content.
         *
         * When working with HTML-formatted bodies, it's important to note that the client may modify the value passed to `prependAsync` in order to
         * make it render efficiently with its rendering engine. This means that the value returned from a subsequent call to the `Body.getAsync` method
         * (introduced in Mailbox 1.3) will not necessarily exactly contain the value that was passed in the `prependAsync` method previously.
         *
         * When including links in HTML markup, you can disable online link preview by setting the `id` attribute on the anchor (\<a\>) to "LPNoLP"
         * (see the **Examples** section for a sample).
         *
         * **Recommended**: Call `getTypeAsync` then pass the returned value to the `options.coercionType` parameter.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `DataExceedsMaximumSize`: The data parameter is longer than 1,000,000 characters.
         *
         * @param data - The string to be inserted at the beginning of the body. The string is limited to 1,000,000 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`. Any errors encountered will be provided in the `asyncResult.error` property.
         */
        prependAsync(data: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Prepends HTML or plain text to the beginning of a message or appointment body when the mail item is sent.
         * 
         * To use `prependOnSendAsync`, you must specify a supplementary permission in the manifest. Details vary with the type of manifest. For guidance,
         * see {@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Understanding Outlook add-in permissions}.
         * 
         * @remarks
         * [Api set: Mailbox preview]
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * **Recommended**: Call `getTypeAsync`, then pass its returned value to the `options.coercionType` parameter.
         * 
         * **Important**: When implementing `prependOnSendAsync`, keep the following in mind.
         * 
         * - In a {@link https://learn.microsoft.com/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough | Smart Alerts add-in},
         * the prepend-on-send feature runs first.
         * 
         * - A new line is added after the prepended content.
         * 
         * - If multiple active add-ins call `prependOnSendAsync`, the order of the inserted content depends on the order in which the add-in runs.
         * The content of the last run add-in appears above previously prepended content.
         * 
         * - If the add-in attempts to insert HTML into a plain text body, the content won't be prepended. Conversely, plain text will be inserted into an HTML body.
         * 
         * **Errors**:
         * 
         * - `DataExceedsMaximumSize`: The `data` parameter exceeds 5,000 characters.
         * 
         * - `InvalidFormatError`: The `options.coercionType` parameter is set to `Office.CoercionType.Html`, but the item body is in plain text format.
         * 
         * @param data - The string to be prepended to the beginning of the message or appointment body. The string is limited to 5,000 characters.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Any object that can be accessed in the callback function.
         *        `coercionType`: The desired format for the body. The string in the `data` parameter is converted to this format.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`. Any errors encountered will be provided in the `asyncResult.error` property.
         * 
         * @beta
         */
        prependOnSendAsync(data: string, options: CommonAPI.AsyncContextOptions & CoercionTypeOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Prepends HTML or plain text to the beginning of a message or appointment body when the mail item is sent.
         * 
         * To use `prependOnSendAsync`, you must specify a supplementary permission in the manifest. Details vary with the type of manifest. For guidance,
         * see {@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Understanding Outlook add-in permissions}.
         * 
         * @remarks
         * [Api set: Mailbox preview]
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * **Recommended**: Call `getTypeAsync`, then pass its returned value to the `options.coercionType` parameter.
         * 
         * **Important**: When implementing `prependOnSendAsync`, keep the following in mind.
         * 
         * - In a {@link https://learn.microsoft.com/office/dev/add-ins/outlook/smart-alerts-onmessagesend-walkthrough | Smart Alerts add-in},
         * the prepend-on-send feature runs first.
         * 
         * - A new line is added after the prepended content.
         * 
         * - If multiple active add-ins call `prependOnSendAsync`, the order of the inserted content depends on the order in which the add-in runs.
         * The content of the last run add-in appears above previously prepended content.
         * 
         * - If the add-in attempts to insert HTML into a plain text body, the content won't be prepended. Conversely, plain text will be inserted into an HTML body.
         * 
         * **Errors**:
         * 
         * - `DataExceedsMaximumSize`: The `data` parameter exceeds 5,000 characters.
         * 
         * - `InvalidFormatError`: The `options.coercionType` parameter is set to `Office.CoercionType.Html`, but the item body is in plain text format.
         * 
         * @param data - The string to be prepended to the beginning of the message or appointment body. The string is limited to 5,000 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`. Any errors encountered will be provided in the `asyncResult.error` property.
         * 
         * @beta
         */
        prependOnSendAsync(data: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Replaces the entire body with the specified text.
         *
         * When working with HTML-formatted bodies, it is important to note that the `Body.getAsync` and `Body.setAsync` methods are not idempotent.
         * The value returned from the `getAsync` method will not necessarily be exactly the same as the value that was passed in the `setAsync` method
         * previously. The client may modify the value passed to `setAsync` in order to make it render efficiently with its rendering engine.
         *
         * When including links in HTML markup, you can disable online link preview by setting the `id` attribute on the anchor (\<a\>) to "LPNoLP"
         * (see the **Examples** section for a sample).
         *
         * **Recommended**: Call `getTypeAsync` then pass the returned value to the `options.coercionType` parameter.
         *
         * **Important**: In Outlook on Windows and on Mac, the add-in user won't be able to revert this action with the **Undo** command.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `DataExceedsMaximumSize`: The data parameter is longer than 1,000,000 characters.
         *
         * - `InvalidFormatError`: The `options.coercionType` parameter is set to `Office.CoercionType.Html` and the message body is in plain text.
         *
         * @param data - The string that will replace the existing body. The string is limited to 1,000,000 characters.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `coercionType`: The desired format for the body. The string in the `data` parameter will be converted to this format.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type Office.AsyncResult. Any errors encountered will be provided in the `asyncResult.error` property.
         */
        setAsync(data: string, options: CommonAPI.AsyncContextOptions & CoercionTypeOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Replaces the entire body with the specified text.
         *
         * When working with HTML-formatted bodies, it is important to note that the `Body.getAsync` and `Body.setAsync` methods are not idempotent.
         * The value returned from the `getAsync` method will not necessarily be exactly the same as the value that was passed in the `setAsync` method
         * previously. The client may modify the value passed to `setAsync` in order to make it render efficiently with its rendering engine.
         *
         * When including links in HTML markup, you can disable online link preview by setting the `id` attribute on the anchor (\<a\>) to "LPNoLP"
         * (see the **Examples** section for a sample).
         *
         * **Recommended**: Call `getTypeAsync` then pass the returned value to the `options.coercionType` parameter.
         *
         * **Important**: In Outlook on Windows and on Mac, the add-in user won't be able to revert this action with the **Undo** command.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `DataExceedsMaximumSize`: The data parameter is longer than 1,000,000 characters.
         *
         * - `InvalidFormatError`: The `options.coercionType` parameter is set to `Office.CoercionType.Html` and the message body is in plain text.
         *
         * @param data - The string that will replace the existing body. The string is limited to 1,000,000 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type Office.AsyncResult. Any errors encountered will be provided in the `asyncResult.error` property.
         */
        setAsync(data: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Replaces the selection in the body with the specified text.
         *
         * The `setSelectedDataAsync` method inserts the specified string at the cursor location in the body of the item, or, if text is selected in
         * the editor, it replaces the selected text. If the cursor was never in the body of the item, or if the body of the item lost focus in the
         * UI, the string will be inserted at the top of the body content. After insertion, the cursor is placed at the end of the inserted content.
         *
         * When including links in HTML markup, you can disable online link preview by setting the `id` attribute on the anchor (\<a\>) to "LPNoLP"
         * (see the **Examples** section for a sample).
         *
         * **Recommended**: Call `getTypeAsync` then pass the returned value to the `options.coercionType` parameter.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `DataExceedsMaximumSize`: The `data` parameter is longer than 1,000,000 characters.
         *
         * - `InvalidFormatError`: The `options.coercionType` parameter is set to `Office.CoercionType.Html` and the message body is in plain text.
         *
         * @param data - The string that will replace the existing body. The string is limited to 1,000,000 characters.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `coercionType`: The desired format for the body. The string in the `data` parameter will be converted to this format.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`. Any errors encountered will be provided in the `asyncResult.error` property.
         */
        setSelectedDataAsync(data: string, options: CommonAPI.AsyncContextOptions & CoercionTypeOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Replaces the selection in the body with the specified text.
         *
         * The `setSelectedDataAsync` method inserts the specified string at the cursor location in the body of the item, or, if text is selected in
         * the editor, it replaces the selected text. If the cursor was never in the body of the item, or if the body of the item lost focus in the
         * UI, the string will be inserted at the top of the body content. After insertion, the cursor is placed at the end of the inserted content.
         *
         * When including links in HTML markup, you can disable online link preview by setting the `id` attribute on the anchor (\<a\>) to "LPNoLP"
         * (see the **Examples** section for a sample).
         *
         * **Recommended**: Call `getTypeAsync` then pass the returned value to the `options.coercionType` parameter.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `DataExceedsMaximumSize`: The `data` parameter is longer than 1,000,000 characters.
         *
         * - `InvalidFormatError`: The `options.coercionType` parameter is set to `Office.CoercionType.Html` and the message body is in plain text.
         *
         * @param data - The string that will replace the existing body. The string is limited to 1,000,000 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`. Any errors encountered will be provided in the `asyncResult.error` property.
         */
        setSelectedDataAsync(data: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Adds or replaces the signature of the item body.
         *
         * **Important**: In Outlook on the web, `setSignatureAsync` only works on messages.
         *
         * **Important**: If your add-in implements the 
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/autolaunch | event-based activation feature using `LaunchEvent` in the manifest},
         * and calls `setSignatureAsync` in the event handler, the following behavior applies.
         *
         * - When the user composes a new item (including reply or forward), the signature is set but doesn't modify the form. This means
         * if the user closes the form without making other edits, they won't be prompted to save changes.
         *
         * @remarks
         * [Api set: Mailbox 1.10]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `DataExceedsMaximumSize`: The `data` parameter is longer than 30,000 characters.
         *
         * - `InvalidFormatError`: The `options.coercionType` parameter is set to `Office.CoercionType.Html` and the message body is in plain text.
         *
         * @param data - The string that represents the signature to be set in the body of the mail. This string is limited to 30,000 characters.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `coercionType`: The format the signature should be set to. If Text, the method sets the signature to plain text,
         *                        removing any HTML tags present. If Html, the method sets the signature to HTML.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`.
         */
        setSignatureAsync(data: string, options: CommonAPI.AsyncContextOptions & CoercionTypeOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Adds or replaces the signature of the item body.
         *
         * **Important**: In Outlook on the web, `setSignatureAsync` only works on messages.
         *
         * **Important**: If your add-in implements the 
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/autolaunch | event-based activation feature using `LaunchEvent` in the manifest},
         * and calls `setSignatureAsync` in the event handler, the following behavior applies.
         *
         * - When the user composes a new item (including reply or forward), the signature is set but doesn't modify the form. This means
         * if the user closes the form without making other edits, they won't be prompted to save changes.
         *
         * @remarks
         * [Api set: Mailbox 1.10]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `DataExceedsMaximumSize`: The `data` parameter is longer than 30,000 characters.
         *
         * - `InvalidFormatError`: The `options.coercionType` parameter is set to `Office.CoercionType.Html` and the message body is in plain text.
         *
         * @param data - The string that represents the signature to be set in the body of the mail. This string is limited to 30,000 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type `Office.AsyncResult`.
         */
        setSignatureAsync(data: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * Represents the categories on an item.
     *
     * In Outlook, a user can tag messages and appointments by using a category to color-code them.
     * The user defines {@link Office.MasterCategories | categories in a master list} on their mailbox.
     * They can then apply one or more categories to an item.
     *
     * **Important**: In Outlook on the web, you can't use the API to manage categories applied to a message in Compose mode.
     *
     * @remarks
     * [Api set: Mailbox 1.8]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface Categories {
        /**
         * Adds categories to an item. Each category must be in the categories master list on that mailbox and so must have a unique name
         * but multiple categories can use the same color.
         *
         * **Important**: In Outlook on the web, you can't use the API to manage categories applied to a message or appointment item in Compose mode.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * **Errors**:
         *
         * - `InvalidCategory`: Invalid categories were provided.
         *
         * @param categories - The categories to be added to the item.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        addAsync(categories: string[], options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Adds categories to an item. Each category must be in the categories master list on that mailbox and so must have a unique name
         * but multiple categories can use the same color.
         *
         * **Important**: In Outlook on the web, you can't use the API to manage categories applied to a message or appointment item in Compose mode.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * **Errors**:
         *
         * - `InvalidCategory`: Invalid categories were provided.
         *
         * @param categories - The categories to be added to the item.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        addAsync(categories: string[], callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Gets an item's categories.
         *
         * **Important**:
         *
         * - If there are no categories on the item, `null` or an empty array will be returned depending on the Outlook version
         * so make sure to handle both cases.
         *
         * - In Outlook on the web, you can't use the API to manage categories applied to a message in Compose mode.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If getting categories fails, the `asyncResult.error` property will contain an error code.
         */
        getAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<CategoryDetails[]>) => void): void;
        /**
         * Gets an item's categories.
         *
         * **Important**:
         *
         * - If there are no categories on the item, `null` or an empty array will be returned depending on the Outlook version
         * so make sure to handle both cases.
         *
         * - In Outlook on the web, you can't use the API to manage categories applied to a message in Compose mode.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If getting categories fails, the `asyncResult.error` property will contain an error code.
         */
        getAsync(callback: (asyncResult: CommonAPI.AsyncResult<CategoryDetails[]>) => void): void;
        /**
         * Removes categories from an item.
         *
         * **Important**: In Outlook on the web, you can't use the API to manage categories applied to a message in Compose mode.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param categories - The categories to be removed from the item.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If removing categories fails, the `asyncResult.error` property will contain an error code.
         */
        removeAsync(categories: string[], options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Removes categories from an item.
         *
         * **Important**: In Outlook on the web, you can't use the API to manage categories applied to a message in Compose mode.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param categories - The categories to be removed from the item.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If removing categories fails, the `asyncResult.error` property will contain an error code.
         */
        removeAsync(categories: string[], callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * Represents a category's details like name and associated color.
     *
     * @remarks
     * [Api set: Mailbox 1.8]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface CategoryDetails {
        /**
         * The name of the category. Maximum length is 255 characters.
         */
        displayName: string;
        /**
         * The color of the category.
         */
        color: MailboxEnums.CategoryColor | string;
    }
    /**
     * Represents the details about a contact (similar to what's on a physical contact or business card) extracted from the item's body. Read mode only.
     *
     * The list of contacts extracted from the body of an email message or appointment is returned in the `contacts` property of the
     * {@link Office.Entities | Entities} object returned by the `getEntities` or `getEntitiesByType` method of the current item.
     *
     * @remarks
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
     */
    export interface Contact {
        /**
         * An array of strings containing the mailing and street addresses associated with the contact. Nullable.
         */
        addresses: string[];
        /**
         * A string containing the name of the business associated with the contact. Nullable.
         */
        businessName: string;
        /**
         * An array of strings containing the SMTP email addresses associated with the contact. Nullable.
         */
        emailAddresses: string[];
        /**
         * A string containing the name of the person associated with the contact. Nullable.
         */
        personName: string;
        /**
         * An array containing a `PhoneNumber` object for each phone number associated with the contact. Nullable.
         */
        phoneNumbers: PhoneNumber[];
        /**
         * An array of strings containing the Internet URLs associated with the contact. Nullable.
         */
        urls: string[];
    }
    /**
     * The `CustomProperties` object represents custom properties that are specific to a particular mail item and specific to an Outlook add-in.
     * For example, there might be a need for an add-in to save some data that's specific to the current message that activated the add-in.
     * If the user revisits the same message in the future and activates the add-in again, the add-in will be able to retrieve the data that had
     * been saved as custom properties. 
     *
     * To learn more about `CustomProperties`, see
     * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/metadata-for-an-outlook-add-in | Get and set add-in metadata for an Outlook add-in}.
     *
     * @remarks
     * [Api set: Mailbox 1.1]
     * 
     * When using custom properties in your add-in, keep in mind that:
     * 
     * - Custom properties saved while in compose mode aren't transmitted to recipients of the mail item. When a message or appointment with custom
     * properties is sent, its properties can be accessed from the item in the Sent Items folder.
     * If you want to make custom data accessible to recipients, consider using {@link Office.InternetHeaders | InternetHeaders} instead.
     * 
     * - The maximum length of a `CustomProperties` JSON object is 2500 characters.
     *
     * - Outlook on Mac doesn't cache custom properties. If the user's network goes down, mail add-ins can't access their custom properties.
     * 
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface CustomProperties {
        /**
         * Returns the value of the specified custom property.
         *
         * @returns The value of the specified custom property.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param name - The name of the custom property to be returned.
         */
        get(name: string): any;
        /**
         * Returns an object with all custom properties in a collection of name/value pairs. The following are equivalent.
         *
         * `customProps.get("name")`
         *
         * `var dictionary = customProps.getAll(); dictionary["name"]`
         *
         * You can iterate through the dictionary object to discover all `names` and `values`.
         *
         * @returns An object with all custom properties in a collection of name/value pairs.
         *
         * @remarks
         * [Api set: Mailbox 1.9]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        getAll(): any;
        /**
         * Removes the specified property from the custom property collection.
         *
         * To make the removal of the property permanent, you must call the `saveAsync` method of the `CustomProperties` object.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param name - The `name` of the property to be removed.
         */
        remove(name: string): void;
        /**
         * Saves custom properties to a message or appointment.
         *
         * You must call the `saveAsync` method to persist any changes made with the `set` method or the `remove` method of the `CustomProperties` object.
         * The saving action is asynchronous.
         *
         * It's a good practice to have your callback function check for and handle errors from `saveAsync`.
         * In particular, a read add-in can be activated while the user is in a connected state in a read form, and subsequently the user becomes
         * disconnected.
         * If the add-in calls `saveAsync` while in the disconnected state, `saveAsync` would return an error.
         * Your callback function should handle this error accordingly.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **Important**: In Outlook on Windows, custom properties saved while in compose mode only persist after the item being composed is closed or
         * after `Office.context.mailbox.item.saveAsync` is called.
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         * @param asyncContext - Optional. Any state data that is passed to the callback function.
         */
        saveAsync(callback: (asyncResult: CommonAPI.AsyncResult<void>) => void, asyncContext?: any): void;
        /**
         * Saves custom properties to a message or appointment.
         *
         * You must call the `saveAsync` method to persist any changes made with the `set` method or the `remove` method of the `CustomProperties` object.
         * The saving action is asynchronous.
         *
         * It's a good practice to have your callback function check for and handle errors from `saveAsync`.
         * In particular, a read add-in can be activated while the user is in a connected state in a read form, and subsequently the user becomes
         * disconnected.
         * If the add-in calls `saveAsync` while in the disconnected state, `saveAsync` would return an error.
         * Your callback function should handle this error accordingly.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param asyncContext - Optional. Any state data that is passed to the callback function.
         */
        saveAsync(asyncContext?: any): void;
        /**
         * Sets the specified property to the specified value.
         *
         * The `set` method sets the specified property to the specified value. To ensure that the set property and value persist on the mail item,
         * you must call the `saveAsync` method.
         *
         * The `set` method creates a new property if the specified property does not already exist;
         * otherwise, the existing value is replaced with the new value.
         * The `value` parameter can be of any type; however, it is always passed to the server as a string.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param name - The name of the property to be set.
         * @param value - The value of the property to be set.
         */
        set(name: string, value: string): void;
    }
    /**
     * The `DelayDeliveryTime` object enables you to manage the delayed delivery date and time of a message.
     *
     * @remarks
     * [Api set: Mailbox preview]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
     *
     * @beta
     */
    export interface DelayDeliveryTime {
        /**
         * Gets the delivery date and time of a message.
         *
         * @remarks
         * [Api set: Mailbox preview]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         *
         * @beta
         */
        getAsync(options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<Date>) => void): void;
        /**
         * Gets the delivery date and time of a message.
         *
         * @remarks
         * [Api set: Mailbox preview]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         *
         * @beta
         */
        getAsync(callback?: (asyncResult: CommonAPI.AsyncResult<Date>) => void): void;
        /**
         * Sets the delivery date and time of a message.
         *
         * @remarks
         * [Api set: Mailbox preview]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param datetime - The future date and time when the message should be sent.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type Office.AsyncResult. Any errors encountered will be provided in the `asyncResult.error` property.
         *
         * @beta
         */
        setAsync(datetime: Date, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Sets the delivery date and time of a message.
         *
         * @remarks
         * [Api set: Mailbox preview]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param datetime - The future date and time when the message should be sent.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                             of type Office.AsyncResult. Any errors encountered will be provided in the `asyncResult.error` property.
         *
         * @beta
         */
        setAsync(datetime: Date, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * Provides diagnostic information to an Outlook add-in.
     *
     * @remarks
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     * 
     * Starting with Mailbox requirement set 1.5, you can also use the 
     * {@link https://learn.microsoft.com/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true#office-office-context-diagnostics-member | Office.context.diagnostics}
     * property to get similar information.
     */
    export interface Diagnostics {
        /**
         * Gets a string that represents the name of the Office application.
         *
         * A string that can be one of the following values: `Outlook`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.
         *
         * **Note**: The `Outlook` value is returned for Outlook on desktop clients (i.e., Windows and Mac).
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        hostName: string;
        /**
         * Gets a string that represents the version of either the Office application or the Exchange Server (e.g., "15.0.468.0").
         *
         * If the mail add-in is running in Outlook on a desktop or mobile client, the `hostVersion` property returns the version of the
         * application, Outlook. In Outlook on the web, the property returns the version of the Exchange Server.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        hostVersion: string;
        /**
         * Gets a string that represents the current view of Outlook on the web.
         *
         * The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.
         *
         * If the application is not Outlook on the web, then accessing this property results in undefined.
         *
         * Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:
         *
         * - `OneColumn`, which is displayed when the screen is narrow. Outlook on the web uses this single-column layout on the entire screen of a
         * smartphone.
         *
         * - `TwoColumns`, which is displayed when the screen is wider. Outlook on the web uses this view on most tablets.
         *
         * - `ThreeColumns`, which is displayed when the screen is wide. For example, Outlook on the web uses this view in a full screen window on a
         * desktop computer.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        OWAView: MailboxEnums.OWAView | "OneColumn" | "TwoColumns" | "ThreeColumns";
    }
    /**
     * Provides the email properties of the sender or specified recipients of an email message or appointment.
     *
     * @remarks
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
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
         * Gets the response that an attendee returned for an appointment.
         * This property applies to only an attendee of an appointment, as represented by the `optionalAttendees` or `requiredAttendees` property.
         * This property returns undefined in other scenarios.
         */
        appointmentResponse: MailboxEnums.ResponseType | string;
        /**
         * Gets the email address type of a recipient.
         * 
         * @remarks
         * **Important**: A `recipientType` property value isn't returned by the 
         * {@link https://learn.microsoft.com/javascript/api/outlook/office.from?view=outlook-js-preview#outlook-office-from-getasync-member(1) | Office.context.mailbox.item.from.getAsync}
         * and {@link https://learn.microsoft.com/javascript/api/outlook/office.organizer?view=outlook-js-preview#outlook-office-organizer-getasync-member(1) | Office.context.mailbox.item.organizer.getAsync} methods.
         * The email sender or appointment organizer is always a user whose email address is on the Exchange server.
         */
        recipientType: MailboxEnums.RecipientType | string;
    }
    /**
     * Represents an email account on an Exchange Server.
     *
     * `EmailUser` objects are primarily received in {@link Office.MeetingSuggestion | MeetingSuggestion} and
     * {@link Office.TaskSuggestion | TaskSuggestion} entities extracted from an Outlook item. To learn more about this scenario, refer to
     * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/extract-entity-strings-from-an-item | Extract entity strings from an Outlook item}.
     *
     * @remarks
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
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
     * Represents the set of locations on an appointment.
     *
     * @remarks
     * [Api set: Mailbox 1.8]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface EnhancedLocation {
        /**
         * Adds to the set of locations associated with the appointment.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `InvalidFormatError`: The format of the specified data object is not valid.
         *
         * @param locationIdentifiers - The locations to be added to the current list of locations.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object. Check the `status` property of `asyncResult` to determine if the call succeeded.
         */
        addAsync(locationIdentifiers: LocationIdentifier[], options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Adds to the set of locations associated with the appointment.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `InvalidFormatError`: The format of the specified data object is not valid.
         *
         * @param locationIdentifiers - The locations to be added to the current list of locations.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object. Check the `status` property of `asyncResult` to determine if the call succeeded.
         */
        addAsync(locationIdentifiers: LocationIdentifier[], callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Gets the set of locations associated with the appointment.
         * 
         * **Note**: {@link https://support.microsoft.com/office/88ff6c60-0a1d-4b54-8c9d-9e1a71bc3023 | Personal contact groups}
         * added as appointment locations aren't returned by this method.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        getAsync(options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<LocationDetails[]>) => void): void;
        /**
         * Gets the set of locations associated with the appointment.
         *
         * **Note**: {@link https://support.microsoft.com/office/88ff6c60-0a1d-4b54-8c9d-9e1a71bc3023 | Personal contact groups}
         * added as appointment locations aren't returned by this method.
         * 
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        getAsync(callback?: (asyncResult: CommonAPI.AsyncResult<LocationDetails[]>) => void): void;
        /**
         * Removes the set of locations associated with the appointment.
         *
         * If there are multiple locations with the same name, all matching locations will be removed even if only one was specified in `locationIdentifiers`.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param locationIdentifiers - The locations to be removed from the current list of locations.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object. Check the `status` property of `asyncResult` to determine if the call succeeded.
         */
        removeAsync(locationIdentifiers: LocationIdentifier[], options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Removes the set of locations associated with the appointment.
         *
         * If there are multiple locations with the same name, all matching locations will be removed even if only one was specified in `locationIdentifiers`.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param locationIdentifiers - The locations to be removed from the current list of locations.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object. Check the `status` property of `asyncResult` to determine if the call succeeded.
         */
        removeAsync(locationIdentifiers: LocationIdentifier[], callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * Provides the current enhanced locations when the `Office.EventType.EnhancedLocationsChanged` event is raised.
     *
     * @remarks
     * [Api set: Mailbox 1.8]
     */
    export interface EnhancedLocationsChangedEventArgs {
        /**
         * Gets the set of enhanced locations.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         */
        enhancedLocations: LocationDetails[];
        /**
         * Gets the type of the event. For details, refer to {@link https://learn.microsoft.com/javascript/api/office/office.eventtype | Office.EventType}.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         */
        type: "olkEnhancedLocationsChanged";
    }
    /**
     * Represents a collection of entities found in an email message or appointment. Read mode only.
     *
     * The `Entities` object is a container for the entity arrays returned by the `getEntities` and `getEntitiesByType` methods when the item
     * (either an email message or an appointment) contains one or more entities that have been found by the server.
     * You can use these entities in your code to provide additional context information to the viewer, such as a map to an address found in the item,
     * or to open a dialer for a phone number found in the item.
     *
     * If no entities of the type specified in the property are present in the item, the property associated with that entity is null.
     * For example, if a message contains a street address and a phone number, the addresses property and phoneNumbers property would contain
     * information, and the other properties would be null.
     *
     * To be recognized as an address, the string must contain a United States postal address that has at least a subset of the elements of a street
     * number, street name, city, state, and zip code.
     *
     * To be recognized as a phone number, the string must contain a North American phone number format.
     *
     * Entity recognition relies on natural language recognition that is based on machine learning of large amounts of data.
     * The recognition of an entity is non-deterministic and success sometimes relies on the particular context in the item.
     *
     * When the property arrays are returned by the `getEntitiesByType` method, only the property for the specified entity contains data;
     * all other properties are null.
     *
     * @remarks
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
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
     * Provides a method to get the from value of a message in an Outlook add-in.
     *
     * @remarks
     * [Api set: Mailbox 1.7]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
     */
    export interface From {
        /**
         * Gets the from value of a message.
         *
         * The `getAsync` method starts an asynchronous call to the Exchange server to get the from value of a message.
         *
         * The from value of the item is provided as an {@link Office.EmailAddressDetails | EmailAddressDetails} in the `asyncResult.value` property.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Important**: A `recipientType` property value isn't returned by the getAsync method.
         * The email sender is always a user whose email address is on the Exchange server.
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                             `asyncResult`, which is an `Office.AsyncResult` object.
         *                  The `value` property of the result is the item's from value, as an `EmailAddressDetails` object.
         */
        getAsync(options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<EmailAddressDetails>) => void): void;
        /**
         * Gets the from value of a message.
         *
         * The `getAsync` method starts an asynchronous call to the Exchange server to get the from value of a message.
         *
         * The from value of the item is provided as an {@link Office.EmailAddressDetails | EmailAddressDetails} in the `asyncResult.value` property.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * **Important**: A `recipientType` property value isn't returned by the getAsync method.
         * The email sender is always a user whose email address is on the Exchange server.
         *
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                             `asyncResult`, which is an `Office.AsyncResult` object.
         *                  The `value` property of the result is the item's from value, as an `EmailAddressDetails` object.
         */
        getAsync(callback?: (asyncResult: CommonAPI.AsyncResult<EmailAddressDetails>) => void): void;
    }
    /**
     * Provides basic details about the notification message that raised the `Office.EventType.InfobarClicked` event.
     *
     * @remarks
     * [Api set: Mailbox 1.10]
     */
    export interface InfobarClickedEventArgs {
        /**
         * Gets additional details about the notification message.
         *
         * @remarks
         * [Api set: Mailbox 1.10]
         */
        infobarDetails: InfobarDetails;
        /**
         * Gets the type of the event. For details, refer to {@link https://learn.microsoft.com/javascript/api/office/office.eventtype | Office.EventType}.
         *
         * @remarks
         * [Api set: Mailbox 1.10]
         */
        type: "olkInfobarClicked";
    }
    /**
     * Provides additional details about the notification message that raised the `Office.EventType.InfobarClicked` event.
     *
     * @remarks
     * [Api set: Mailbox 1.10]
     */
    export interface InfobarDetails {
        /**
         * The action type. Currently, "Dismiss" is the only supported action.
         *
         * @remarks
         * [Api set: Mailbox 1.10]
         */
        actionType: MailboxEnums.InfobarActionType;
        /**
         * The notification type.
         *
         * @remarks
         * [Api set: Mailbox 1.10]
         */
        infobarType: MailboxEnums.InfobarType;
    }
    /**
     * The `InternetHeaders` object represents custom internet headers that are preserved after the message item leaves Exchange
     * and is converted to a MIME message. These headers are stored as x-headers in the MIME message.
     *
     * Internet headers are stored as string key-value pairs on a per-item basis.
     *
     * **Note**: This object is intended for you to set and get your custom headers on a message item. To learn more, see
     * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/internet-headers | Get and set internet headers on a message in an Outlook add-in}.
     *
     * @remarks
     * [Api set: Mailbox 1.8]
     *
     * **Recommended practices**
     *
     * Currently, internet headers are a finite resource on a user's mailbox. When the quota is exhausted, you can't create any more internet headers
     * on that mailbox, which can result in unexpected behavior from clients that rely on this to function.
     *
     * Apply the following guidelines when you create internet headers in your add-in.
     *
     * - Create the minimum number of headers required. The header quota is based on the total size of headers applied to a message. In Exchange Online,
     * the header limit is capped at 256 KB, while in an Exchange on-premises environment, the limit is determined by your organization's administrator.
     * For further information on header limits, see 
     * {@link https://learn.microsoft.com/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits#message-limits | Exchange Online message limits}
     * and {@link https://learn.microsoft.com/exchange/mail-flow/message-size-limits?view=exchserver-2019#types-of-message-size-limits | Exchange Server message limits}.
     *
     * - Name headers so that you can reuse and update their values later. As such, avoid naming headers in a variable manner
     * (for example, based on user input, timestamp, etc.).
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
     */
    export interface InternetHeaders {
        /**
         * Given an array of internet header names, this method returns a record containing those internet headers and their values.
         * If the add-in requests an x-header that isn't available, that x-header will not be returned in the results.
         *
         * **Note**: This method is intended to return the values of the custom headers you set using the `setAsync` method.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param names - The names of the internet headers to be returned.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, of type `Office.AsyncResult`. The string key-value pairs of internet headers are returned in the
         *                `asyncResult.value` property. Any errors encountered are provided in the `asyncResult.error` property.
         */
        getAsync(names: string[], options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<Record<string, string>>) => void): void;
        /**
         * Given an array of internet header names, this method returns a record containing those internet headers and their values.
         * If the add-in requests an x-header that isn't available, that x-header will not be returned in the results.
         *
         * **Note**: This method is intended to return the values of the custom headers you set using the `setAsync` method.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param names - The names of the internet headers to be returned.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, of type `Office.AsyncResult`. The string key-value pairs of internet headers are returned in the
         *                `asyncResult.value` property. Any errors encountered are provided in the `asyncResult.error` property.
         */
        getAsync(names: string[], callback: (asyncResult: CommonAPI.AsyncResult<Record<string, string>>) => void): void;
        /**
         * Given an array of internet header names, this method removes the specified headers from the internet header collection.
         *
         * **Note**: This method is intended to remove the custom headers you set using the `setAsync` method.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param names - The names of the internet headers to be removed.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, of type `Office.AsyncResult`. Any errors encountered are provided in the `asyncResult.error` property.
         */
        removeAsync(names: string[], options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Given an array of internet header names, this method removes the specified headers from the internet header collection.
         *
         * **Note**: This method is intended to remove the custom headers you set using the `setAsync` method.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param names - The names of the internet headers to be removed.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, of type `Office.AsyncResult`. Any errors encountered are provided in the `asyncResult.error` property.
         */
        removeAsync(names: string[], callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Sets the specified internet headers to the specified values.
         *
         * The `setAsync` method creates a new header if the specified header doesn't already exist; otherwise, the existing value is replaced with
         * the new value.
         *
         * **Note**: This method is intended to set the values of your custom headers.
         * 
         * **Important**: The header quota is based on the total size of headers applied to a message. In Exchange Online,
         * the header limit is capped at 256 KB, while in an Exchange on-premises environment, the limit is determined by your organization's administrator.
         * For further information on header limits, see 
         * {@link https://learn.microsoft.com/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits#message-limits | Exchange Online message limits}
         * and {@link https://learn.microsoft.com/exchange/mail-flow/message-size-limits?view=exchserver-2019#types-of-message-size-limits | Exchange Server message limits}.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param headers - The names and corresponding values of the headers to be set. This should be a record object with its keys being internet header names
         *                and values being the corresponding header value strings.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                 `asyncResult`, of type `Office.AsyncResult`. Any errors encountered are provided in the `asyncResult.error` property.
         */
        setAsync(headers: Record<string, string>, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Sets the specified internet headers to the specified values.
         *
         * The `setAsync` method creates a new header if the specified header doesn't already exist; otherwise, the existing value is replaced with
         * the new value.
         *
         * **Note**: This method is intended to set the values of your custom headers.
         * 
         * **Important**: The header quota is based on the total size of headers applied to a message. In Exchange Online,
         * the header limit is capped at 256 KB, while in an Exchange on-premises environment, the limit is determined by your organization's administrator.
         * For further information on header limits, see 
         * {@link https://learn.microsoft.com/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits#message-limits | Exchange Online message limits}
         * and {@link https://learn.microsoft.com/exchange/mail-flow/message-size-limits?view=exchserver-2019#types-of-message-size-limits | Exchange Server message limits}.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param headers - The names and corresponding values of the headers to be set. This should be a record object with its keys being internet header names
         *                and values being the corresponding header value strings.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                 `asyncResult`, of type `Office.AsyncResult`. Any errors encountered are provided in the `asyncResult.error` property.
         */
        setAsync(headers: Record<string, string>, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * The item namespace is used to access the currently selected message, meeting request, or appointment.
     * You can determine the type of the item by using the `itemType` property.
     *
     * To see the full member list, refer to the
     * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item | Object Model} page.
     *
     * If you want to see IntelliSense for only a specific type or mode, cast this item to one of the following:
     *
     * - {@link Office.AppointmentCompose | AppointmentCompose}
     *
     * - {@link Office.AppointmentRead | AppointmentRead}
     *
     * - {@link Office.MessageCompose | MessageCompose}
     *
     * - {@link Office.MessageRead | MessageRead}
     *
     * @remarks
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Appointment Organizer, Appointment Attendee, Message Compose, Message Read
     */
    export interface Item {
    }
    /**
     * The compose mode of {@link Office.Item | Office.context.mailbox.item}.
     *
     * **Important**: This is an internal Outlook object, not directly exposed through existing interfaces.
     * You should treat this as a mode of `Office.context.mailbox.item`. For more information, refer to the
     * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item | Object Model} page.
     *
     * Child interfaces:
     *
     * - {@link Office.AppointmentCompose | AppointmentCompose}
     *
     * - {@link Office.MessageCompose | MessageCompose}
     */
    export interface ItemCompose extends Item {
    }
    /**
     * The read mode of {@link Office.Item | Office.context.mailbox.item}.
     *
     * **Important**: This is an internal Outlook object, not directly exposed through existing interfaces.
     * You should treat this as a mode of `Office.context.mailbox.item`. For more information, refer to the
     * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item | Object Model} page.
     *
     * Child interfaces:
     *
     * - {@link Office.AppointmentRead | AppointmentRead}
     *
     * - {@link Office.MessageRead | MessageRead}
     */
    export interface ItemRead extends Item {
    }
    /**
     * Represents a date and time in the local client's time zone. Read mode only.
     *
     * @remarks
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
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
         * Integer value representing the year.
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
     * @remarks
     * [Api set: Mailbox 1.1]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
     */
    export interface Location {
        /**
         * Gets the location of an appointment.
         *
         * The `getAsync` method starts an asynchronous call to the Exchange server to get the location of an appointment.
         * The location of the appointment is provided as a string in the `asyncResult.value` property.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        getAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Gets the location of an appointment.
         *
         * The `getAsync` method starts an asynchronous call to the Exchange server to get the location of an appointment.
         * The location of the appointment is provided as a string in the `asyncResult.value` property.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        getAsync(callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Sets the location of an appointment.
         *
         * The `setAsync` method starts an asynchronous call to the Exchange server to set the location of an appointment.
         * Setting the location of an appointment overwrites the current location.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - DataExceedsMaximumSize: The location parameter is longer than 255 characters.
         *
         * @param location - The location of the appointment. The string is limited to 255 characters.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If setting the location fails, the `asyncResult.error` property will contain an error code.
         */
        setAsync(location: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Sets the location of an appointment.
         *
         * The `setAsync` method starts an asynchronous call to the Exchange server to set the location of an appointment.
         * Setting the location of an appointment overwrites the current location.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - DataExceedsMaximumSize: The location parameter is longer than 255 characters.
         *
         * @param location - The location of the appointment. The string is limited to 255 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If setting the location fails, the `asyncResult.error` property will contain an error code.
         */
        setAsync(location: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * Represents a location. Read-only.
     *
     * @remarks
     * [Api set: Mailbox 1.8]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface LocationDetails {
        /**
         * The `LocationIdentifier` of the location.
         */
        locationIdentifier: LocationIdentifier;
        /**
         * The location's display name.
         */
        displayName: string;
        /**
         * The email address associated with the location. Only locations of type `Room` have an email address.
         */
        emailAddress: string;
    }
    /**
     * Represents the ID of a location.
     *
     * @remarks
     * [Api set: Mailbox 1.8]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface LocationIdentifier {
        /**
         * The location's unique ID.
         *
         * For `Room` type, it's the room's email address.
         *
         * For `Custom` type, it's the `displayName`.
         */
        id: string;
        /**
         * The location's type.
         */
        type: MailboxEnums.LocationType | string;
    }
    /**
     * Provides access to the Microsoft Outlook add-in object model.
     *
     * Key properties:
     *
     * - `diagnostics`: Provides diagnostic information to an Outlook add-in.
     *
     * - `item`: Provides methods and properties for accessing a message or appointment in an Outlook add-in.
     *
     * - `userProfile`: Provides information about the user in an Outlook add-in.
     *
     * @remarks
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface Mailbox {
        /**
         * Provides diagnostic information to an Outlook add-in.
         *
         * Contains the following members.
         *
         *  - `hostName` (string): A string that represents the name of the Office application.
         * It should be one of the following values: `Outlook`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.
         * **Note**: The "Outlook" value is returned for Outlook on desktop clients (i.e., Windows and Mac).
         *
         *  - `hostVersion` (string): A string that represents the version of either the Office application or the Exchange Server (e.g., "15.0.468.0").
         * If the mail add-in is running in Outlook on desktop or mobile clients, the `hostVersion` property returns the version of the
         * application, Outlook. In Outlook on the web, the property returns the version of the Exchange Server.
         *
         *  - `OWAView` (`MailboxEnums.OWAView` or string): An enum (or string literal) that represents the current view of Outlook on the web.
         * If the application is not Outlook on the web, then accessing this property results in undefined.
         * Outlook on the web has three views (`OneColumn` - displayed when the screen is narrow, `TwoColumns` - displayed when the screen is wider,
         * and `ThreeColumns` - displayed when the screen is wide) that correspond to the width of the screen and the window, and the number of columns
         * that can be displayed.
         *
         * More information is under {@link Office.Diagnostics}.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         * 
         * Starting with Mailbox requirement set 1.5, you can also use the 
         * {@link https://learn.microsoft.com/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true#office-office-context-diagnostics-member | Office.context.diagnostics}
         * property to get similar information.
         */
        diagnostics: Diagnostics;
        /**
         * Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.
         *
         * Your app must have the **read item** permission specified in its manifest to call the `ewsUrl` member in read mode.
         *
         * In compose mode you must call the `saveAsync` method before you can use the `ewsUrl` member.
         * Your app must have **read/write item** permissions to call the `saveAsync` method.
         *
         * **Note**: This member is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox.
         * For example, you can create a remote service to {@link https://learn.microsoft.com/office/dev/add-ins/outlook/get-attachments-of-an-outlook-item | get attachments from the selected item}.
         */
        ewsUrl: string;
        /**
         * The mailbox item. Depending on the context in which the add-in opened, the item type may vary.
         * If you want to see IntelliSense for only a specific type or mode, cast this item to one of the following:
         *
         * {@link Office.MessageCompose | MessageCompose}, {@link Office.MessageRead | MessageRead},
         * {@link Office.AppointmentCompose | AppointmentCompose}, {@link Office.AppointmentRead | AppointmentRead}
         *
         * **Important**: `item` can be null if your add-in supports pinning the task pane. For details on how to handle, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/pinnable-taskpane#implement-the-event-handler | Implement a pinnable task pane in Outlook}.
         */
        item?: Item & ItemCompose & ItemRead & Message & MessageCompose & MessageRead & Appointment & AppointmentCompose & AppointmentRead;
        /**
         * Gets an object that provides methods to manage the categories master list associated with a mailbox.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write mailbox**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        masterCategories: MasterCategories;
        /**
         * Gets the URL of the REST endpoint for this email account.
         *
         * Your app must have the **read item** permission specified in its manifest to call the `restUrl` member in read mode.
         *
         * In compose mode you must call the `saveAsync` method before you can use the `restUrl` member.
         * Your app must have **read/write item** permissions to call the `saveAsync` method.
         *
         * However, in delegate or shared scenarios, you should instead use the `targetRestUrl` property of the
         * {@link Office.SharedProperties | SharedProperties} object (introduced in requirement set 1.8). For more information,
         * see the {@link https://learn.microsoft.com/office/dev/add-ins/outlook/delegate-access | shared folders and shared mailbox} article.
         *
         * @remarks
         * [Api set: Mailbox 1.5]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * The `restUrl` value can be used to make {@link https://learn.microsoft.com/outlook/rest/ | REST API} calls to the user's mailbox.
         */
        restUrl: string;
        /**
         * Information about the user associated with the mailbox. This includes their account type, display name, email address, and time zone.
         *
         * More information is under {@link Office.UserProfile}
         */
        userProfile: UserProfile;

        /**
         * Adds an event handler for a supported event. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Mailbox object model
         * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.5]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param eventType - The event that should invoke the handler.
         * @param handler - The function to handle the event. The function must accept a single parameter, which is an object literal.
         *                The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        addHandlerAsync(eventType: CommonAPI.EventType | string, handler: any, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Adds an event handler for a supported event. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Mailbox object model
         * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.5]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param eventType - The event that should invoke the handler.
         * @param handler - The function to handle the event. The function must accept a single parameter, which is an object literal.
         *                The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        addHandlerAsync(eventType: CommonAPI.EventType | string, handler: any, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Converts an item ID formatted for REST into EWS format.
         *
         * Item IDs retrieved via a REST API (such as the Outlook Mail API or the Microsoft Graph) use a different format than the format used by
         * Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param itemId - An item ID formatted for the Outlook REST APIs.
         * @param restVersion - A value indicating the version of the Outlook REST API used to retrieve the item ID.
         */
        convertToEwsId(itemId: string, restVersion: MailboxEnums.RestVersion | string): string;
        /**
         * Gets a dictionary containing time information in local client time.
         *
         * The dates and times used by a mail app for Outlook on the web or desktop clients can use different time zones.
         * Outlook uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).
         * You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that
         * the user expects.
         *
         * If the mail app is running in Outlook on desktop clients, the `convertToLocalClientTime` method will return a dictionary object
         * with the values set to the client computer time zone.
         * If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object
         * with the values set to the time zone specified in the EAC.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param timeValue - A `Date` object.
         */
        convertToLocalClientTime(timeValue: Date): LocalClientTime;
        /**
         * Converts an item ID formatted for EWS into REST format.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the
         * {@link https://learn.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations | Outlook Mail API}
         * or the {@link https://graph.microsoft.io/ | Microsoft Graph}.
         * The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.
         *
         * @param itemId - An item ID formatted for Exchange Web Services (EWS)
         * @param restVersion - A value indicating the version of the Outlook REST API that the converted ID will be used with.
         */
        convertToRestId(itemId: string, restVersion: MailboxEnums.RestVersion | string): string;
        /**
         * Gets a `Date` object from a dictionary containing time information.
         *
         * The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a `Date` object with the correct values for the
         * local date and time.
         *
         * @returns A Date object with the time expressed in UTC.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param input - The local time value to convert.
         */
        convertToUtcClientTime(input: LocalClientTime): Date;
        /**
         * Displays an existing calendar appointment.
         *
         * The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on
         * mobile devices.
         *
         * In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment
         * of a recurring series. However, you can't display an instance of the series because you can't access the properties
         * (including the item ID) of instances of a recurring series.
         *
         * In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32K characters.
         *
         * If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and
         * no error message is returned.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param itemId - The Exchange Web Services (EWS) identifier for an existing calendar appointment.
         */
        displayAppointmentForm(itemId: string): void;
        /**
         * Displays an existing calendar appointment.
         *
         * The `displayAppointmentFormAsync` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on
         * mobile devices.
         *
         * In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment
         * of a recurring series. However, you can't display an instance of the series because you can't access the properties
         * (including the item ID) of instances of a recurring series.
         *
         * In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32K characters.
         *
         * If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and
         * no error message is returned.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.9]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param itemId - The Exchange Web Services (EWS) identifier for an existing calendar appointment.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        displayAppointmentFormAsync(itemId: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Displays an existing calendar appointment.
         *
         * The `displayAppointmentFormAsync` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on
         * mobile devices.
         *
         * In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment
         * of a recurring series. However, you can't display an instance of the series because you can't access the properties
         * (including the item ID) of instances of a recurring series.
         *
         * In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32K characters.
         *
         * If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and
         * no error message is returned.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.9]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param itemId - The Exchange Web Services (EWS) identifier for an existing calendar appointment.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        displayAppointmentFormAsync(itemId: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Displays an existing message.
         *
         * The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.
         *
         * In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32K characters.
         *
         * If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and
         * no error message is returned.
         *
         * Do not use the `displayMessageForm` with an itemId that represents an appointment. Use the `displayAppointmentForm` method to display
         * an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param itemId - The Exchange Web Services (EWS) identifier for an existing message.
         */
        displayMessageForm(itemId: string): void;
        /**
         * Displays an existing message.
         *
         * The `displayMessageFormAsync` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.
         *
         * In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32K characters.
         *
         * If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and
         * no error message is returned.
         *
         * Do not use the `displayMessageForm` or `displayMessageFormAsync` method with an itemId that represents an appointment.
         * Use the `displayAppointmentForm` or `displayAppointmentFormAsync` method to display an existing appointment,
         * and `displayNewAppointmentForm` or `displayNewAppointmentFormAsync` to display a form to create a new appointment.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.9]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param itemId - The Exchange Web Services (EWS) identifier for an existing message.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        displayMessageFormAsync(itemId: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Displays an existing message.
         *
         * The `displayMessageFormAsync` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.
         *
         * In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32K characters.
         *
         * If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and
         * no error message is returned.
         *
         * Do not use the `displayMessageForm` or `displayMessageFormAsync` method with an itemId that represents an appointment.
         * Use the `displayAppointmentForm` or `displayAppointmentFormAsync` method to display an existing appointment,
         * and `displayNewAppointmentForm` or `displayNewAppointmentFormAsync` to display a form to create a new appointment.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.9]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param itemId - The Exchange Web Services (EWS) identifier for an existing message.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        displayMessageFormAsync(itemId: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Displays a form for creating a new calendar appointment.
         *
         * The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting.
         * If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.
         *
         * In Outlook on the web, this method always displays a form with an attendees field.
         * If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.
         * If you have specified attendees, the form would include the attendees and a **Send** button.
         *
         * In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or
         * `resources` parameter, this method displays a meeting form with a **Send** button.
         * If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.
         *
         * If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
         *
         * @param parameters - An `AppointmentForm` describing the new appointment. All properties are optional.
         */
        displayNewAppointmentForm(parameters: AppointmentForm): void;
        /**
         * Displays a form for creating a new calendar appointment.
         *
         * The `displayNewAppointmentFormAsync` method opens a form that enables the user to create a new appointment or meeting.
         * If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.
         *
         * In Outlook on the web, this method always displays a form with an attendees field.
         * If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.
         * If you have specified attendees, the form would include the attendees and a **Send** button.
         *
         * In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or
         * `resources` parameter, this method displays a meeting form with a **Send** button.
         * If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.
         *
         * If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.9]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
         *
         * @param parameters - An `AppointmentForm` describing the new appointment. All properties are optional.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        displayNewAppointmentFormAsync(parameters: AppointmentForm, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Displays a form for creating a new calendar appointment.
         *
         * The `displayNewAppointmentFormAsync` method opens a form that enables the user to create a new appointment or meeting.
         * If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.
         *
         * In Outlook on the web, this method always displays a form with an attendees field.
         * If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.
         * If you have specified attendees, the form would include the attendees and a **Send** button.
         *
         * In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or
         * `resources` parameter, this method displays a meeting form with a **Send** button.
         * If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.
         *
         * If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.9]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
         *
         * @param parameters - An `AppointmentForm` describing the new appointment. All properties are optional.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        displayNewAppointmentFormAsync(parameters: AppointmentForm, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Displays a form for creating a new message.
         *
         * The `displayNewMessageForm` method opens a form that enables the user to create a new message. If parameters are specified, the message form
         * fields are automatically populated with the contents of the parameters.
         *
         * If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.
         *
         * @remarks
         * [Api set: Mailbox 1.6]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
         *
         * @param parameters - A dictionary containing all values to be filled in for the user in the new form. All parameters are optional.
         *
         *        `toRecipients`: An array of strings containing the email addresses or an array containing an {@link Office.EmailAddressDetails | EmailAddressDetails} object
         *        for each of the recipients on the **To** line. The array is limited to a maximum of 100 entries.
         *
         *        `ccRecipients`: An array of strings containing the email addresses or an array containing an {@link Office.EmailAddressDetails | EmailAddressDetails} object
         *        for each of the recipients on the **Cc** line. The array is limited to a maximum of 100 entries.
         *
         *        `bccRecipients`: An array of strings containing the email addresses or an array containing an {@link Office.EmailAddressDetails | EmailAddressDetails} object
         *        for each of the recipients on the **Bcc** line. The array is limited to a maximum of 100 entries.
         *
         *        `subject`: A string containing the subject of the message. The string is limited to a maximum of 255 characters.
         *
         *        `htmlBody`: The HTML body of the message. The body content is limited to a maximum size of 32 KB.
         *
         *        `attachments`: An array of JSON objects that are either file or item attachments.
         *
         *        `attachments.type`: Indicates the type of attachment. Must be file for a file attachment or item for an item attachment.
         *
         *        `attachments.name`: A string that contains the name of the attachment, up to 255 characters in length.
         *
         *        `attachments.url`: Only used if type is set to file. The URI of the location for the file. **Important**: This link must be
         *        publicly accessible, without need for authentication by Exchange Online servers. However, with on-premises Exchange, the link can
         *        be accessible on a private network as long as it doesn't need further authentication.
         *
         *        `attachments.isInline`: Only used if type is set to file. If true, indicates that the attachment will be shown inline in the
         *        message body, and should not be displayed in the attachment list.
         *
         *        `attachments.itemId`: Only used if type is set to item. The EWS item ID of the existing e-mail you want to attach to the new message.
         *        This is a string up to 100 characters.
         */
        displayNewMessageForm(parameters: any): void;
        /**
         * Displays a form for creating a new message.
         *
         * The `displayNewMessageFormAsync` method opens a form that enables the user to create a new message.
         * If parameters are specified, the message form fields are automatically populated with the contents of the parameters.
         *
         * If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.
         *
         * @remarks
         * [Api set: Mailbox 1.9]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
         *
         * @param parameters - A dictionary containing all values to be filled in for the user in the new form. All parameters are optional.
         *
         *        `toRecipients`: An array of strings containing the email addresses or an array containing an {@link Office.EmailAddressDetails | EmailAddressDetails} object
         *        for each of the recipients on the **To** line. The array is limited to a maximum of 100 entries.
         *
         *        `ccRecipients`: An array of strings containing the email addresses or an array containing an {@link Office.EmailAddressDetails | EmailAddressDetails} object
         *        for each of the recipients on the **Cc** line. The array is limited to a maximum of 100 entries.
         *
         *        `bccRecipients`: An array of strings containing the email addresses or an array containing an {@link Office.EmailAddressDetails | EmailAddressDetails} object
         *        for each of the recipients on the **Bcc** line. The array is limited to a maximum of 100 entries.
         *
         *        `subject`: A string containing the subject of the message. The string is limited to a maximum of 255 characters.
         *
         *        `htmlBody`: The HTML body of the message. The body content is limited to a maximum size of 32 KB.
         *
         *        `attachments`: An array of JSON objects that are either file or item attachments.
         *
         *        `attachments.type`: Indicates the type of attachment. Must be file for a file attachment or item for an item attachment.
         *
         *        `attachments.name`: A string that contains the name of the attachment, up to 255 characters in length.
         *
         *        `attachments.url`: Only used if type is set to file. The URI of the location for the file. **Important**: This link must be
         *        publicly accessible, without need for authentication by Exchange Online servers. However, with on-premises Exchange, the link can
         *        be accessible on a private network as long as it doesn't need further authentication.
         *
         *        `attachments.isInline`: Only used if type is set to file. If true, indicates that the attachment will be shown inline in the
         *        message body, and should not be displayed in the attachment list.
         *
         *        `attachments.itemId`: Only used if type is set to item. The EWS item ID of the existing e-mail you want to attach to the new message.
         *        This is a string up to 100 characters.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        displayNewMessageFormAsync(parameters: any, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Displays a form for creating a new message.
         *
         * The `displayNewMessageFormAsync` method opens a form that enables the user to create a new message.
         * If parameters are specified, the message form fields are automatically populated with the contents of the parameters.
         *
         * If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.
         *
         * @remarks
         * [Api set: Mailbox 1.9]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
         *
         * @param parameters - A dictionary containing all values to be filled in for the user in the new form. All parameters are optional.
         *
         *        `toRecipients`: An array of strings containing the email addresses or an array containing an {@link Office.EmailAddressDetails | EmailAddressDetails} object
         *        for each of the recipients on the **To** line. The array is limited to a maximum of 100 entries.
         *
         *        `ccRecipients`: An array of strings containing the email addresses or an array containing an {@link Office.EmailAddressDetails | EmailAddressDetails} object
         *        for each of the recipients on the **Cc** line. The array is limited to a maximum of 100 entries.
         *
         *        `bccRecipients`: An array of strings containing the email addresses or an array containing an {@link Office.EmailAddressDetails | EmailAddressDetails} object
         *        for each of the recipients on the **Bcc** line. The array is limited to a maximum of 100 entries.
         *
         *        `subject`: A string containing the subject of the message. The string is limited to a maximum of 255 characters.
         *
         *        `htmlBody`: The HTML body of the message. The body content is limited to a maximum size of 32 KB.
         *
         *        `attachments`: An array of JSON objects that are either file or item attachments.
         *
         *        `attachments.type`: Indicates the type of attachment. Must be file for a file attachment or item for an item attachment.
         *
         *        `attachments.name`: A string that contains the name of the attachment, up to 255 characters in length.
         *
         *        `attachments.url`: Only used if type is set to file. The URI of the location for the file. **Important**: This link must be
         *        publicly accessible, without need for authentication by Exchange Online servers. However, with on-premises Exchange, the link can
         *        be accessible on a private network as long as it doesn't need further authentication.
         *
         *        `attachments.isInline`: Only used if type is set to file. If true, indicates that the attachment will be shown inline in the
         *        message body, and should not be displayed in the attachment list.
         *
         *        `attachments.itemId`: Only used if type is set to item. The EWS item ID of the existing e-mail you want to attach to the new message.
         *        This is a string up to 100 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        displayNewMessageFormAsync(parameters: any, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Gets a string that contains a token used to call REST APIs or Exchange Web Services (EWS).
         *
         * The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox.
         * The lifetime of the callback token is 5 minutes.
         *
         * The token is returned as a string in the `asyncResult.value` property.
         *
         * Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **read item**.
         *
         * Calling the `getCallbackTokenAsync` method in compose mode requires you to have saved the item.
         * The `saveAsync` method requires a minimum permission level of **read/write item**.
         *
         * **Important**: For guidance on delegate or shared scenarios, see the
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/delegate-access | shared folders and shared mailbox} article.
         *
         * *REST Tokens*
         *
         * When a REST token is requested (`options.isRest` = `true`), the resulting token will not work to authenticate EWS calls.
         * The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the
         * **read/write mailbox** permission in its manifest.
         * If the **read/write mailbox** permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts,
         * including the ability to send mail.
         *
         * The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.
         *
         * This API works for the following scopes.
         *
         * - `Mail.ReadWrite`
         *
         * - `Mail.Send`
         *
         * - `Calendars.ReadWrite`
         *
         * - `Contacts.ReadWrite`
         *
         * *EWS Tokens*
         *
         * When an EWS token is requested (`options.isRest` = `false`), the resulting token will not work to authenticate REST API calls.
         * The token will be limited in scope to accessing the current item.
         *
         * The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.
         *
         * You can pass both the token and either an attachment identifier or item identifier to an external system. That system uses
         * the token as a bearer authorization token to call the Exchange Web Services (EWS)
         * {@link https://learn.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation | GetAttachment} operation or
         * {@link https://learn.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation | GetItem} operation to return an
         * attachment or item. For example, you can create a remote service to
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/get-attachments-of-an-outlook-item | get attachments from the selected item}.
         *
         * **Important**: EWS operations aren't supported in add-ins running in Outlook on iOS and Android. A REST token is always returned in Outlook
         * mobile clients even if `options.isRest` is set to `false`.
         *
         * **Note**: It's recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.
         *
         * @remarks
         * [Api set: Mailbox 1.5]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * **Errors**:
         *
         * - `HTTPRequestFailure`: The request has failed. Please look at the diagnostics object for the HTTP error code.
         *
         * - `InternalServerError`: The Exchange server returned an error. Please look at the diagnostics object for more information.
         *
         * - `NetworkError`: The user is no longer connected to the network. Please check your network connection and try again.
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `isRest`: Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.
         *        `asyncContext`: Any state data that is passed to the asynchronous method.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. The token is returned as a string in the `asyncResult.value` property.
         *                 If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.
         */
        getCallbackTokenAsync(options: CommonAPI.AsyncContextOptions & { isRest?: boolean }, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Gets a string that contains a token used to get an attachment or item from an Exchange Server.
         *
         * The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox.
         * The lifetime of the callback token is 5 minutes.
         *
         * The token is returned as a string in the `asyncResult.value` property.
         *
         * You can pass both the token and either an attachment identifier or item identifier to an external system. That system uses
         * the token as a bearer authorization token to call the Exchange Web Services (EWS)
         * {@link https://learn.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation | GetAttachment} or
         * {@link https://learn.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation | GetItem} operation to return an
         * attachment or item. For example, you can create a remote service to
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/get-attachments-of-an-outlook-item | get attachments from the selected item}.
         *
         * Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **read item**.
         *
         * Calling the `getCallbackTokenAsync` method in compose mode requires you to have saved the item.
         * The `saveAsync` method requires a minimum permission level of **read/write item**.
         *
         * **Important**: For guidance on delegate or shared scenarios, see the
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/delegate-access | shared folders and shared mailbox} article.
         *
         * **Note**: This method isn't supported in Outlook on iOS or Android. EWS operations aren't supported in add-ins running on Outlook mobile clients.
         * 
         * @remarks
         * [Api set: All support Read mode; Mailbox 1.3 introduced Compose mode support]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * **Errors**:
         *
         * - `HTTPRequestFailure`: The request has failed. Please look at the diagnostics object for the HTTP error code.
         *
         * - `InternalServerError`: The Exchange server returned an error. Please look at the diagnostics object for more information.
         *
         * - `NetworkError`: The user is no longer connected to the network. Please check your network connection and try again.
         *
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. The token is returned as a string in the `asyncResult.value` property.
         *                 If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.
         * @param userContext - Optional. Any state data that is passed to the asynchronous method.
         */
        getCallbackTokenAsync(callback: (asyncResult: CommonAPI.AsyncResult<string>) => void, userContext?: any): void;
        /**
         * Gets currently selected messages on which an add-in can activate and perform operations. An add-in can activate on a maximum of 100 messages at a time.
         * To learn more about item multi-select, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/item-multi-select | Activate your Outlook add-in on multiple messages (preview)}.
         *
         * @remarks
         * [Api set: Mailbox preview]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write mailbox**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose, Read
         * 
         * **Important**: This method only applies to messages.
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *        `asyncResult`, which is an `Office.AsyncResult` object. The properties of the selected messages are returned as an array of JSON objects in the
         *        `asyncResult.value` property. These properties include the item ID, subject, item type (`Message` is the only supported type at this time), and item mode
         *        (`Read` or `Compose`). The objects in the array follow the order in which messages were selected.
         * 
         * @beta
         */
        getSelectedItemsAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<object[]>) => void): void;
        /**
         * Gets currently selected messages on which an add-in can activate and perform operations. An add-in can activate on a maximum of 100 messages at a time.
         * To learn more about item multi-select, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/item-multi-select | Activate your Outlook add-in on multiple messages (preview)}.
         *
         * @remarks
         * [Api set: Mailbox preview]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write mailbox**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose, Read
         * 
         * **Important**: This method only applies to messages.
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *        `asyncResult`, which is an `Office.AsyncResult` object. The properties of the selected messages are returned as an array of JSON objects in the
         *        `asyncResult.value` property. These properties include the item ID, subject, item type (`Message` is the only supported type at this time), and item mode
         *        (`Read` or `Compose`). The objects in the array follow the order in which messages were selected.
         * 
         * @beta
         */
        getSelectedItemsAsync(callback: (asyncResult: CommonAPI.AsyncResult<object[]>) => void): void;
        /**
         * Gets a token identifying the user and the Office Add-in.
         *
         * The token is returned as a string in the `asyncResult.value` property.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * The `getUserIdentityTokenAsync` method returns a token that you can use to identify and
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/authentication | authenticate the add-in and user with an external system}.
         *
         * **Errors**:
         *
         * - `HTTPRequestFailure`: The request has failed. Please look at the diagnostics object for the HTTP error code.
         *
         * - `InternalServerError`: The Exchange server returned an error. Please look at the diagnostics object for more information.
         *
         * - `NetworkError`: The user is no longer connected to the network. Please check your network connection and try again.
         *
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         *                 The token is returned as a string in the `asyncResult.value` property.
         *                 If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.
         * @param userContext - Optional. Any state data that is passed to the asynchronous method.
         */
        getUserIdentityTokenAsync(callback: (asyncResult: CommonAPI.AsyncResult<string>) => void, userContext?: any): void;
        /**
         * Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user's mailbox.
         *
         * In these cases, add-ins should use REST APIs to access the user's mailbox instead.
         *
         * The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.
         *
         * You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.
         *
         * The XML request must specify UTF-8 encoding: `\<?xml version="1.0" encoding="utf-8"?\>`.
         *
         * Your add-in must have the **read/write mailbox** permission to use the `makeEwsRequestAsync` method.
         * For information about using the **read/write mailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method,
         * see {@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Specify permissions for mail add-in access to the user's mailbox}.
         *
         * The XML result of the EWS call is provided as a string in the `asyncResult.value` property.
         * If the result exceeds 1 MB in size, an error message is returned instead.
         *
         * **Note**: This method is not supported in the following scenarios.
         *
         * - In Outlook on iOS or Android.
         *
         * - When the add-in is loaded in a Gmail mailbox.
         *
         * **Note**: The server administrator must set `OAuthAuthentication` to `true` on the Client Access Server EWS directory to enable the
         * `makeEwsRequestAsync` method to make EWS requests.
         *
         * *Version differences*
         *
         * When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set
         * the encoding value to ISO-8859-1.
         *
         * `<?xml version="1.0" encoding="iso-8859-1"?>`
         *
         * You do not need to set the encoding value when your mail app is running in Outlook on the web.
         * You can determine whether your mail app is running in Outlook or Outlook on the web by using the `mailbox.diagnostics.hostName` property.
         * You can determine what version of Outlook is running by using the `mailbox.diagnostics.hostVersion` property.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write mailbox**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param data - The EWS request.
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter
         *                   of type `Office.AsyncResult`.
         *                 The `value` property of the result is the XML of the EWS request provided as a string.
         *                 If the result exceeds 1 MB in size, an error message is returned instead.
         * @param userContext - Optional. Any state data that is passed to the asynchronous method.
         */
        makeEwsRequestAsync(data: any, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void, userContext?: any): void;
        /**
         * Removes the event handlers for a supported event type. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Mailbox object model
         * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.5]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param eventType - The event that should revoke the handler.
         * @param options - Provides an option for preserving context data of any type, unchanged, for use in a callback.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        removeHandlerAsync(eventType: CommonAPI.EventType | string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Removes the event handlers for a supported event type. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Mailbox object model
         * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.5]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param eventType - The event that should revoke the handler.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        removeHandlerAsync(eventType: CommonAPI.EventType | string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * Represents the categories master list on the mailbox.
     *
     * In Outlook, a user can tag messages and appointments by using a category to color-code them.
     * The user defines categories in a master list on their mailbox. They can then apply one or more categories to an item.
     *
     * **Important**: In delegate or shared scenarios, the delegate can get the categories in the master list but can't add or remove categories.
     *
     * @remarks
     * [Api set: Mailbox 1.8]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write mailbox**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface MasterCategories {
        /**
         * Adds categories to the master list on a mailbox. Each category must have a unique name but multiple categories can use the same color.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write mailbox**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * **Errors**:
         *
         * - `DuplicateCategory`: One of the categories provided is already in the master category list.
         *
         * - `PermissionDenied`: The user does not have permission to perform this action.
         *
         * @param categories - The categories to be added to the master list on the mailbox.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        addAsync(categories: CategoryDetails[], options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Adds categories to the master list on a mailbox. Each category must have a unique name but multiple categories can use the same color.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write mailbox**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * **Errors**:
         *
         * - `DuplicateCategory`: One of the categories provided is already in the master category list.
         *
         * - `PermissionDenied`: The user does not have permission to perform this action.
         *
         * @param categories - The categories to be added to the master list on the mailbox.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        addAsync(categories: CategoryDetails[], callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Gets the master list of categories on a mailbox.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write mailbox**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If adding categories fails, the `asyncResult.error` property will contain an error code.
         */
        getAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<CategoryDetails[]>) => void): void;
        /**
         * Gets the master list of categories on a mailbox.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write mailbox**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        getAsync(callback: (asyncResult: CommonAPI.AsyncResult<CategoryDetails[]>) => void): void;
        /**
         * Removes categories from the master list on a mailbox.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write mailbox**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * **Errors**:
         *
         * - `PermissionDenied`: The user does not have permission to perform this action.
         *
         * @param categories - The categories to be removed from the master list on the mailbox.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If removing categories fails, the `asyncResult.error` property will contain an error code.
         */
        removeAsync(categories: string[], options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Removes categories from the master list on a mailbox.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write mailbox**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * **Errors**:
         *
         * - `PermissionDenied`: The user does not have permission to perform this action.
         *
         * @param categories - The categories to be removed from the master list on the mailbox.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If removing categories fails, the `asyncResult.error` property will contain an error code.
         */
        removeAsync(categories: string[], callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * Represents a suggested meeting found in an item. Read mode only.
     *
     * The list of meetings suggested in an email message is returned in the `meetingSuggestions` property of the `Entities` object that is returned when
     * the `getEntities` or `getEntitiesByType` method is called on the active item.
     *
     * The start and end values are string representations of a `Date` object that contains the date and time at which the suggested meeting is to
     * begin and end. The values are in the default time zone specified for the current user.
     *
     * @remarks
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
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
        meetingString: string;
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
     * A subclass of {@link Office.Item | Item} for messages.
     *
     * **Important**: This is an internal Outlook object, not directly exposed through existing interfaces.
     * You should treat this as a mode of `Office.context.mailbox.item`. For more information, refer to the
     * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item | Object Model} page.
     *
     * Child interfaces:
     *
     * - {@link Office.MessageCompose | MessageCompose}
     *
     * - {@link Office.MessageRead | MessageRead}
     */
    export interface Message extends Item {
    }
     /**
     * The message compose mode of {@link Office.Item | Office.context.mailbox.item}.
     *
     * **Important**: This is an internal Outlook object, not directly exposed through existing interfaces.
     * You should treat this as a mode of `Office.context.mailbox.item`. For more information, refer to the
     * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item | Object Model} page.
     *
     * Parent interfaces:
     *
     * - {@link Office.ItemCompose | ItemCompose}
     *
     * - {@link Office.Message | Message}
     */
    export interface MessageCompose extends Message, ItemCompose {
        /**
         * Gets an object that provides methods to get or update the recipients on the **Bcc** (blind carbon copy) line of a message.
         *
         * Depending on the client/platform (i.e., Windows, Mac, etc.), limits may apply on how many recipients you can get or update.
         * See the {@link Office.Recipients | Recipients} object for more details.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        bcc: Recipients;
        /**
         * Gets an object that provides methods for manipulating the body of an item.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        body: Body;
        /**
         * Gets an object that provides methods for managing the item's categories.
         *
         * **Important**: In Outlook on the web, you can't use the API to manage categories on a message in Compose mode.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        categories: Categories;
        /**
         * Provides access to the Cc (carbon copy) recipients of a message. The type of object and level of access depend on the mode of the
         * current item.
         *
         * The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the
         * **Cc** line of the message. However, depending on the client/platform (i.e., Windows, Mac, etc.), limits may apply on how many recipients
         * you can get or update. See the {@link Office.Recipients | Recipients} object for more details.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        cc: Recipients;
        /**
         * Gets an identifier for the email conversation that contains a particular message.
         *
         * You can get an integer for this property if your mail app is activated in read forms or responses in compose forms.
         * If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change
         * and that value you obtained earlier will no longer apply.
         *
         * You get null for this property for a new item in a compose form.
         * If the user sets a subject and saves the item, the `conversationId` property will return a value.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        conversationId: string;
        /**
         * Gets or sets the delayed delivery date and time of a message.
         *
         * The `delayDeliveryTime` property returns a `DelayDeliveryTime` object that provides methods to manage the delivery date and time of the message.
         *
         * @remarks
         * [Api set: Mailbox preview]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @beta
         */
        delayDeliveryTime: DelayDeliveryTime;
        /**
         * Gets the email address of the sender of a message.
         *
         * The `from` property returns a `From` object that provides a method to get the from value.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        from: From;
        /**
         * Gets or sets the custom internet headers of a message.
         *
         * The `internetHeaders` property returns an `InternetHeaders` object that provides methods to manage the internet headers on the message.
         *
         * To learn more, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/internet-headers | Get and set internet headers on a message in an Outlook add-in}.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        internetHeaders: InternetHeaders;
        /**
         * Gets the type of item that an instance represents.
         *
         * The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the item object instance is a message or
         * an appointment.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        itemType: MailboxEnums.ItemType | string;
        /**
         * Gets the notification messages for an item.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        notificationMessages: NotificationMessages;
        /**
         * Gets the object to get or set the {@link Office.SensitivityLabel | sensitivity label} of a message.
         * 
         * @remarks
         * [Api set: Mailbox preview]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * **Important**: To use the sensitivity label feature in your add-in, you must have a Microsoft 365 E5 subscription.
         *
         * To learn more about how to manage sensitivity labels in your add-in, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/sensitivity-label | Manage the sensitivity label of your message or appointment in compose mode (preview)}.
         *
         * @beta
         */
        sensitivityLabel: SensitivityLabel;
        /**
         * Gets the ID of the series that an instance belongs to.
         *
         * In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item
         * that this item belongs to. However, on iOS and Android, the seriesId returns the REST ID of the parent item.
         *
         * **Note**: The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.
         * The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.
         * Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`.
         * For more details, see {@link https://learn.microsoft.com/office/dev/add-ins/outlook/use-rest-api | Use the Outlook REST APIs from an Outlook add-in}.
         *
         * The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests
         * and returns `undefined` for any other items that are not meeting requests.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        seriesId: string;
        /**
         * Manages the {@link Office.SessionData | SessionData} of an item in Compose mode.
         *
         * **Important**: The entire SessionData object is limited to 50,000 characters per add-in.
         *
         * @remarks
         * [Api set: Mailbox 1.11]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        sessionData: SessionData;
        /**
         * Gets or sets the description that appears in the subject field of an item.
         *
         * The `subject` property gets or sets the entire subject of the item, as sent by the email server.
         *
         * The `subject` property returns a `Subject` object that provides methods to get and set the subject.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        subject: Subject;
        /**
         * Provides access to the recipients on the **To** line of a message. The type of object and level of access depend on the mode of the
         * current item.
         *
         * The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the
         * **To** line of the message. However, depending on the client/platform (i.e., Windows, Mac, etc.), limits may apply on how many recipients
         * you can get or update. See the {@link Office.Recipients | Recipients} object for more details.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        to: Recipients;

        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * You can subsequently use the identifier with the `removeAttachmentAsync` method to remove the attachment in the same session.
         *
         * **Important**: In recent builds of Outlook on Windows, a bug was introduced that incorrectly appends an `Authorization: Bearer` header to
         * this action (whether using this API or the Outlook UI). To work around this issue, you can try using the `addFileAttachmentFromBase64` API
         * introduced with requirement set 1.8.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * **Errors**:
         *
         * - `AttachmentSizeExceeded`: The attachment is larger than allowed.
         *
         * - `FileTypeNotSupported`: The attachment has an extension that is not allowed.
         *
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         *
         * @param uri - The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `isInline`: If true, indicates that the attachment will be shown inline in the message body, and should not be displayed in the
         *        attachment list.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. On success, the attachment identifier will be provided in the `asyncResult.value` property.
         *                 If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of
         *                 the error.
         */
        addFileAttachmentAsync(uri: string, attachmentName: string, options: CommonAPI.AsyncContextOptions & { isInline: boolean }, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.
         *
         * You can subsequently use the identifier with the `removeAttachmentAsync` method to remove the attachment in the same session.
         *
         * **Important**: In recent builds of Outlook on Windows, a bug was introduced that incorrectly appends an `Authorization: Bearer` header to
         * this action (whether using this API or the Outlook UI). To work around this issue, you can try using the `addFileAttachmentFromBase64` API
         * introduced with requirement set 1.8.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * **Errors**:
         *
         * - `AttachmentSizeExceeded`: The attachment is larger than allowed.
         *
         * - `FileTypeNotSupported`: The attachment has an extension that is not allowed.
         *
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         *
         * @param uri - The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. On success, the attachment identifier will be provided in the `asyncResult.value` property.
         *                 If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of
         *                 the error.
         */
        addFileAttachmentAsync(uri: string, attachmentName: string, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The `addFileAttachmentFromBase64Async` method uploads the file from the Base64 encoding and attaches it to the item in the compose form.
         * This method returns the attachment identifier in the `asyncResult.value` object.
         *
         * You can subsequently use the identifier with the `removeAttachmentAsync` method to remove the attachment in the same session.
         *
         * **Note**: If you're using a data URL API (e.g., `readAsDataURL`), you need to strip out the data URL prefix then send the rest of the string to this API.
         * For example, if the full string is represented by `data:image/svg+xml;base64,<rest of Base64 string>`, remove `data:image/svg+xml;base64,`.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * **Errors**:
         *
         * - `AttachmentSizeExceeded`: The attachment is larger than allowed.
         *
         * - `FileTypeNotSupported`: The attachment has an extension that isn't allowed.
         *
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         *
         * **Note**: If you're adding an inline Base64 image to the body of a message or appointment being composed, you must first get the current item body using the
         * {@link https://learn.microsoft.com/javascript/api/outlook/office.body#outlook-office-body-getasync-member(1) | Office.context.mailbox.item.body.getAsync} 
         * method before inserting the image using `addFileAttachmentFromBase64Async`. Otherwise, the image won't render in the body once it's inserted.
         * For further guidance, see {@link https://learn.microsoft.com/office/dev/add-ins/outlook/add-and-remove-attachments-to-an-item-in-a-compose-form#attach-a-file | Attach a file}.
         *
         * @param base64File - The Base64-encoded content of an image or file to be added to an email or event. The maximum length of the encoded string is 27,892,122 characters (about 25 MB).
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `isInline`: If true, indicates that the attachment will be shown inline in the message body and should not be displayed in the attachment list.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                             type Office.AsyncResult. On success, the attachment identifier will be provided in the `asyncResult.value` property.
         *                  If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.
         */
        addFileAttachmentFromBase64Async(base64File: string, attachmentName: string, options: CommonAPI.AsyncContextOptions & { isInline: boolean }, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Adds a file to a message or appointment as an attachment.
         *
         * The `addFileAttachmentFromBase64Async` method uploads the file from the Base64 encoding and attaches it to the item in the compose form.
         * This method returns the attachment identifier in the `asyncResult.value` object.
         *
         * You can subsequently use the identifier with the `removeAttachmentAsync` method to remove the attachment in the same session.
         *
         * **Note**: If you're using a data URL API (e.g., `readAsDataURL`), you need to strip out the data URL prefix then send the rest of the string to this API.
         * For example, if the full string is represented by `data:image/svg+xml;base64,<rest of Base64 string>`, remove `data:image/svg+xml;base64,`.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * **Errors**:
         *
         * - `AttachmentSizeExceeded`: The attachment is larger than allowed.
         *
         * - `FileTypeNotSupported`: The attachment has an extension that isn't allowed.
         *
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         *
         * **Note**: If you're adding an inline Base64 image to the body of a message or appointment being composed, you must first get the current item body using the
         * {@link https://learn.microsoft.com/javascript/api/outlook/office.body#outlook-office-body-getasync-member(1) | Office.context.mailbox.item.body.getAsync} 
         * method before inserting the image using `addFileAttachmentFromBase64Async`. Otherwise, the image won't render in the body once it's inserted.
         * For further guidance, see {@link https://learn.microsoft.com/office/dev/add-ins/outlook/add-and-remove-attachments-to-an-item-in-a-compose-form#attach-a-file | Attach a file}.
         *
         * @param base64File - The Base64-encoded content of an image or file to be added to an email or event. The maximum length of the encoded string is 27,892,122 characters (about 25 MB).
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                             type Office.AsyncResult. On success, the attachment identifier will be provided in the `asyncResult.value` property.
         *                  If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.
         */
        addFileAttachmentFromBase64Async(base64File: string, attachmentName: string, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Adds an event handler for a supported event. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Item object model
         * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param eventType - The event that should invoke the handler.
         * @param handler - The function to handle the event. The function must accept a single parameter, which is an object literal.
         *                The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        addHandlerAsync(eventType: CommonAPI.EventType | string, handler: any, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Adds an event handler for a supported event. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Item object model
         * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param eventType - The event that should invoke the handler.
         * @param handler - The function to handle the event. The function must accept a single parameter, which is an object literal.
         *                The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        addHandlerAsync(eventType: CommonAPI.EventType | string, handler: any, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message or appointment.
         *
         * The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form.
         * If you specify a callback function, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or
         * a code that indicates any error that occurred while attaching the item. You can use the options parameter to pass state information to the
         * callback function, if needed.
         *
         * You can subsequently use the identifier with the `removeAttachmentAsync` method to remove the attachment in the same session.
         *
         * If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that
         * you are editing; however, this is not supported and is not recommended.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * **Errors**:
         *
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         *
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. On success, the attachment identifier will be provided in the `asyncResult.value` property.
         *                 If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of
         *                 the error.
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Adds an Exchange item, such as a message, as an attachment to the message or appointment.
         *
         * The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form.
         * If you specify a callback function, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or
         * a code that indicates any error that occurred while attaching the item. You can use the options parameter to pass state information to the
         * callback function, if needed.
         *
         * You can subsequently use the identifier with the `removeAttachmentAsync` method to remove the attachment in the same session.
         *
         * If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that
         * you are editing; however, this is not supported and is not recommended.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * **Errors**:
         *
         * - `NumberOfAttachmentsExceeded`: The message or appointment has too many attachments.
         *
         * @param itemId - The Exchange identifier of the item to attach. The maximum length is 100 characters.
         * @param attachmentName - The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. On success, the attachment identifier will be provided in the `asyncResult.value` property.
         *                 If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of
         *                 the error.
         */
        addItemAttachmentAsync(itemId: any, attachmentName: string, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Closes the current item that is being composed
         *
         * The behaviors of the close method depends on the current state of the item being composed.
         * If the item has unsaved changes, the client prompts the user to save, discard, or close the action.
         *
         * In the Outlook desktop client, if the message is an inline reply, the close method has no effect.
         *
         * **Note**: In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save,
         * discard, or cancel even if no changes have occurred since the item was last saved.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         */
        close(): void;
        /**
         * Disables the Outlook client signature.
         *
         * For Windows and Mac rich clients, this API sets the signature under the "New Message" and "Replies/Forwards" sections
         * for the sending account to "(none)", effectively disabling the signature.
         * For Outlook on the web, the API should disable the signature option for new mails, replies, and forwards.
         * If the signature is selected, this API call should disable it.
         *
         * @remarks
         * [Api set: Mailbox 1.10]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        disableClientSignatureAsync(options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Disables the Outlook client signature.
         *
         * For Windows and Mac rich clients, this API sets the signature under the "New Message" and "Replies/Forwards" sections
         * for the sending account to "(none)", effectively disabling the signature.
         * For Outlook on the web, the API should disable the signature option for new mails, replies, and forwards.
         * If the signature is selected, this API call should disable it.
         *
         * @remarks
         * [Api set: Mailbox 1.10]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param callback - Optional. When the method completes, the function passed in the callback parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        disableClientSignatureAsync(callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Gets an attachment from a message or appointment and returns it as an `AttachmentContent` object.
         *
         * The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item. As a best practice, you should get
         * the attachment's identifier from a `getAttachmentsAsync` call, then in the same session, use that identifier to retrieve the attachment.
         * In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.
         * A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to
         * continue in a separate window.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * **Errors**:
         *
         * - `AttachmentTypeNotSupported`: The attachment type isn't supported. Unsupported types include embedded images in Rich Text Format,
         *                               or item attachment types other than email or calendar items (such as a contact or task item).
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param attachmentId - The identifier of the attachment you want to get.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object. If the call fails, the `asyncResult.error` property will contain
         *                an error code with the reason for the failure.
         */
        getAttachmentContentAsync(attachmentId: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<AttachmentContent>) => void): void;
        /**
         * Gets an attachment from a message or appointment and returns it as an `AttachmentContent` object.
         *
         * The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item. As a best practice, you should get
         * the attachment's identifier from a `getAttachmentsAsync` call, then in the same session, use that identifier to retrieve the attachment.
         * In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.
         * A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to
         * continue in a separate window.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * **Errors**:
         *
         * - `AttachmentTypeNotSupported`: The attachment type isn't supported. Unsupported types include embedded images in Rich Text Format,
         *                               or item attachment types other than email or calendar items (such as a contact or task item).
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param attachmentId - The identifier of the attachment you want to get.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object. If the call fails, the `asyncResult.error` property will contain
         *                an error code with the reason for the failure.
         */
        getAttachmentContentAsync(attachmentId: string, callback?: (asyncResult: CommonAPI.AsyncResult<AttachmentContent>) => void): void;
        /**
         * Gets the item's attachments as an array.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If the call fails, the `asyncResult.error` property will contain an error code with the reason for
         *                 the failure.
         */
        getAttachmentsAsync(options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<AttachmentDetailsCompose[]>) => void): void;
        /**
         * Gets the item's attachments as an array.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If the call fails, the `asyncResult.error` property will contain an error code with the reason for
         *                 the failure.
         */
        getAttachmentsAsync(callback?: (asyncResult: CommonAPI.AsyncResult<AttachmentDetailsCompose[]>) => void): void;
        /**
         * Specifies the type of message compose and its coercion type. The message can be new, or a reply or forward.
         * The coercion type can be HTML or plain text.
         *
         * @remarks
         * [Api set: Mailbox 1.10]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. On success, the `asyncResult.value` property contains an object with the item's compose type
         *                 and coercion type.
         *
         * @returns
         * An object with `ComposeType` and `CoercionType` enum values for the message item.
         */
        getComposeTypeAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<any>) => void): void;
        /**
         * Specifies the type of message compose and its coercion type. The message can be new, or a reply or forward.
         * The coercion type can be HTML or plain text.
         *
         * @remarks
         * [Api set: Mailbox 1.10]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. On success, the `asyncResult.value` property contains an object with the item's compose type
         *                 and coercion type.
         *
         * @returns
         * An object with `ComposeType` and `CoercionType` enum values for the message item.
         */
        getComposeTypeAsync(callback: (asyncResult: CommonAPI.AsyncResult<any>) => void): void;
        /**
         * Gets initialization data passed when the add-in is {@link https://learn.microsoft.com/outlook/actionable-messages/invoke-add-in | activated by an actionable message}.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         *                 On success, the initialization context data is provided as a string (or an empty string if there's no initialization context)
         *                 in the `asyncResult.value` property..
         */
        getInitializationContextAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Gets initialization data passed when the add-in is {@link https://learn.microsoft.com/outlook/actionable-messages/invoke-add-in | activated by an actionable message}.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         *                 On success, the initialization context data is provided as a string (or an empty string if there's no initialization context)
         *                 in the `asyncResult.value` property.
         */
        getInitializationContextAsync(callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Asynchronously gets the ID of a saved item.
         *
         * When invoked, this method returns the item ID via the callback function.
         *
         * **Note**: If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API),
         * be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.
         * Until the item is synced, the `itemId` is not recognized and using it returns an error.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * **Errors**:
         *
         * - `ItemNotSaved`: The ID can't be retrieved until the item is saved.
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                   of type `Office.AsyncResult`.
         */
        getItemIdAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Asynchronously gets the ID of a saved item.
         *
         * When invoked, this method returns the item ID via the callback function.
         *
         * **Note**: If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API),
         * be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.
         * Until the item is synced, the `itemId` is not recognized and using it returns an error.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * **Errors**:
         *
         * - `ItemNotSaved`: The ID can't be retrieved until the item is saved.
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                   of type `Office.AsyncResult`.
         */
        getItemIdAsync(callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Asynchronously returns selected data from the subject or body of a message.
         *
         * If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data.
         * If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.
         *
         * To access the selected data from the callback function, call `asyncResult.value.data`.
         * To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.
         *
         * @returns
         * The selected data as a string with format determined by `coercionType`.
         *
         * @remarks
         * [Api set: Mailbox 1.2]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param coercionType - Requests a format for the data. If `Text`, the method returns the plain text as a string, removing any HTML tags present.
         *                     If `Html`, the method returns the selected text, whether it is plaintext or HTML.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        getSelectedDataAsync(coercionType: CommonAPI.CoercionType | string, options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<any>) => void): void;
        /**
         * Asynchronously returns selected data from the subject or body of a message.
         *
         * If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data.
         * If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.
         *
         * To access the selected data from the callback function, call `asyncResult.value.data`.
         * To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.
         *
         * @returns
         * The selected data as a string with format determined by `coercionType`.
         *
         * @remarks
         * [Api set: Mailbox 1.2]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param coercionType - Requests a format for the data. If `Text`, the method returns the plain text as a string, removing any HTML tags present.
         *                     If `Html`, the method returns the selected text, whether it is plaintext or HTML.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        getSelectedDataAsync(coercionType: CommonAPI.CoercionType | string, callback: (asyncResult: CommonAPI.AsyncResult<any>) => void): void;
        /**
         * Gets the properties of an appointment or message in a shared folder or shared mailbox (now in preview).
         *
         * For more information around using this API, see the
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/delegate-access | shared folders and shared mailbox} article.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **Important**: In Message Compose mode, this API is not supported in Outlook on the web or on Windows unless the following conditions are met.
         *
         * a. **Delegate access/Shared folders**
         *
         * 1. The mailbox owner starts a message. This can be a new message, a reply, or a forward.
         *
         * 2. They save the message then move it from their own **Drafts** folder to a folder shared with the delegate.
         *
         * 3. The delegate opens the draft from the shared folder then continues composing.
         *
         * b. **Shared mailbox (applies to Outlook on Windows only)**
         *
         * 1. The shared mailbox user starts a message. This can be a new message, a reply, or a forward.
         *
         * 2. They save the message then move it from their own **Drafts** folder to a folder in the shared mailbox.
         *
         * 3. Another shared mailbox user opens the draft from the shared mailbox then continues composing.
         *
         * The message is now in a shared context and add-ins that support these shared scenarios can get the item's shared properties.
         * After the message has been sent, it's usually found in the sender's **Sent Items** folder.
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. The `value` property of the result is the properties of the shared item.
         */
        getSharedPropertiesAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<SharedProperties>) => void): void;
        /**
         * Gets the properties of an appointment or message in a shared folder or shared mailbox (now in preview).
         *
         * For more information around using this API, see the
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/delegate-access | shared folders and shared mailbox} article.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **Important**: In Message Compose mode, this API is not supported in Outlook on the web or on Windows unless the following conditions are met.
         *
         * a. **Delegate access/Shared folders**
         *
         * 1. The mailbox owner starts a message. This can be a new message, a reply, or a forward.
         *
         * 2. They save the message then move it from their own **Drafts** folder to a folder shared with the delegate.
         *
         * 3. The delegate opens the draft from the shared folder then continues composing.
         *
         * b. **Shared mailbox (applies to Outlook on Windows only)**
         *
         * 1. The shared mailbox user starts a message. This can be a new message, a reply, or a forward.
         *
         * 2. They save the message then move it from their own **Drafts** folder to a folder in the shared mailbox.
         *
         * 3. Another shared mailbox user opens the draft from the shared mailbox then continues composing.
         *
         * The message is now in a shared context and add-ins that support these shared scenarios can get the item's shared properties.
         * After the message has been sent, it's usually found in the sender's **Sent Items** folder.
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. The `value` property of the result is the properties of the shared item.
         */
        getSharedPropertiesAsync(callback: (asyncResult: CommonAPI.AsyncResult<SharedProperties>) => void): void;
        /**
         * Gets if the client signature is enabled.
         *
         * For Windows and Mac rich clients, the API call should return `true` if the default signature for new messages, replies, or forwards is set
         * to a template for the sending Outlook account.
         * For Outlook on the web, the API call should return `true` if the signature is enabled for compose types `newMail`, `reply`, or `forward`.
         * If the settings are set to "(none)" in Mac or Windows rich clients or disabled in Outlook on the Web, the API call should return `false`.
         *
         * @remarks
         * [Api set: Mailbox 1.10]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                   type `Office.AsyncResult`.
         */
        isClientSignatureEnabledAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<boolean>) => void): void;
        /**
         * Gets if the client signature is enabled.
         *
         * For Windows and Mac rich clients, the API call should return `true` if the default signature for new messages, replies, or forwards is set
         * to a template for the sending Outlook account.
         * For Outlook on the web, the API call should return `true` if the signature is enabled for compose types `newMail`, `reply`, or `forward`.
         * If the settings are set to "(none)" in Mac or Windows rich clients or disabled in Outlook on the Web, the API call should return `false`.
         *
         * @remarks
         * [Api set: Mailbox 1.10]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                   type `Office.AsyncResult`.
         */
        isClientSignatureEnabledAsync(callback: (asyncResult: CommonAPI.AsyncResult<boolean>) => void): void;
        /**
         * Asynchronously loads custom properties for this add-in on the selected item.
         *
         * Custom properties are stored as key-value pairs on a per-app, per-item basis.
         * This method returns a {@link Office.CustomProperties | CustomProperties} object in the callback, which provides methods to access the custom properties specific to the
         * current item and the current add-in. Custom properties aren't encrypted on the item, so this shouldn't be used as secure storage.
         *
         * The custom properties are provided as a `CustomProperties` object in the `asyncResult.value` property.
         * This object can be used to get, set, save, and remove custom properties from the mail item.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * To learn more about custom properties, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/metadata-for-an-outlook-add-in | Get and set add-in metadata for an Outlook add-in}.
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         * @param userContext - Optional. Developers can provide any object they wish to access in the callback function.
         *                    This object can be accessed by the `asyncResult.asyncContext` property in the callback function.
         */
        loadCustomPropertiesAsync(callback: (asyncResult: CommonAPI.AsyncResult<CustomProperties>) => void, userContext?: any): void;
        /**
         * Removes an attachment from a message or appointment.
         *
         * The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.
         * As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment
         * in the same session. In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.
         * A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to
         * continue in a separate window.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * **Errors**:
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param attachmentId - The identifier of the attachment to remove. The maximum string length of the `attachmentId`
         *                       is 200 characters in Outlook on the web and on Windows.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If removing the attachment fails, the `asyncResult.error` property will contain an error code
         *                 with the reason for the failure.
         */
        removeAttachmentAsync(attachmentId: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Removes an attachment from a message or appointment.
         *
         * The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.
         * As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment
         * in the same session. In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.
         * A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to
         * continue in a separate window.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * **Errors**:
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param attachmentId - The identifier of the attachment to remove. The maximum string length of the `attachmentId`
         *                       is 200 characters in Outlook on the web and on Windows.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If removing the attachment fails, the `asyncResult.error` property will contain an error code
         *                 with the reason for the failure.
         */
        removeAttachmentAsync(attachmentId: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Removes the event handlers for a supported event type. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Item object model
         * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param eventType - The event that should revoke the handler.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        removeHandlerAsync(eventType: CommonAPI.EventType | string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Removes the event handlers for a supported event type. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Item object model
         * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * @param eventType - The event that should revoke the handler.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        removeHandlerAsync(eventType: CommonAPI.EventType | string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item ID via the callback function.
         * In Outlook on the web or Outlook in online mode, the item is saved to the server.
         * In Outlook in cached mode, the item is saved to the local cache.
         *
         * Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal
         * appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent.
         * Saving an existing appointment will send an update to added or removed attendees.
         *
         * When working with HTML-formatted content, it's important to note that the Outlook client may modify the content. This means that
         * subsequent calls to methods like `Body.getAsync`, `Body.setAsync`, and even `saveAsync` may not result in the same content.
         *
         * **Note**: If your add-in calls `saveAsync` on an item in compose mode in order to get an item ID to use with EWS or the REST API, be aware
         * that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.
         * Until the item is synced, using the itemId will return an error.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * **Errors**:
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        saveAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Asynchronously saves an item.
         *
         * When invoked, this method saves the current message as a draft and returns the item ID via the callback function.
         * In Outlook on the web or Outlook in online mode, the item is saved to the server.
         * In Outlook in cached mode, the item is saved to the local cache.
         *
         * Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal
         * appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent.
         * Saving an existing appointment will send an update to added or removed attendees.
         *
         * When working with HTML-formatted content, it's important to note that the Outlook client may modify the content. This means that
         * subsequent calls to methods like `Body.getAsync`, `Body.setAsync`, and even `saveAsync` may not result in the same content.
         *
         * **Note**: If your add-in calls `saveAsync` on an item in compose mode in order to get an item ID to use with EWS or the REST API, be aware
         * that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.
         * Until the item is synced, using the `itemId` will return an error.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * **Errors**:
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        saveAsync(callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         *
         * The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is
         * selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned.
         * After insertion, the cursor is placed at the end of the inserted content.
         *
         * @remarks
         * [Api set: Mailbox 1.2]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * **Errors**:
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters.
         *             If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         *        `coercionType`: If text, the current style is applied in Outlook on the web and desktop clients.
         *        If the field is an HTML editor, only the text data is inserted, even if the data is HTML.
         *        If html and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is
         *        applied in Outlook on desktop clients. If the field is a text field, an `InvalidDataFormat` error is returned.
         *        If `coercionType` is not set, the result depends on the field:
         *        if the field is HTML then HTML is used; if the field is text, then plain text is used.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        setSelectedDataAsync(data: string, options: CommonAPI.AsyncContextOptions & CoercionTypeOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Asynchronously inserts data into the body or subject of a message.
         *
         * The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is
         * selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned.
         * After insertion, the cursor is placed at the end of the inserted content.
         *
         * @remarks
         * [Api set: Mailbox 1.2]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Compose
         *
         * **Errors**:
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param data - The data to be inserted. Data is not to exceed 1,000,000 characters.
         *             If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        setSelectedDataAsync(data: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * The message read mode of {@link Office.Item | Office.context.mailbox.item}.
     *
     * **Important**: This is an internal Outlook object, not directly exposed through existing interfaces.
     * You should treat this as a mode of `Office.context.mailbox.item`. For more information, refer to the
     * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item | Object Model} page.
     *
     * Parent interfaces:
     *
     * - {@link Office.ItemRead | ItemRead}
     *
     * - {@link Office.Message | Message}
     */
    export interface MessageRead extends Message, ItemRead {
        /**
         * Gets the item's attachments as an array.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * **Note**: Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.
         * For more information, see
         * {@link https://support.microsoft.com/office/434752e1-02d3-4e90-9124-8b81e49a8519 | Blocked attachments in Outlook}.
         *
         */
        attachments: AttachmentDetails[];
        /**
         * Gets an object that provides methods for manipulating the body of an item.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        body: Body;
        /**
         * Gets an object that provides methods for managing the item's categories.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        categories: Categories;
        /**
         * Provides access to the Cc (carbon copy) recipients of a message. The type of object and level of access depend on the mode of the
         * current item.
         *
         * The `cc` property returns an array that contains an {@link Office.EmailAddressDetails | EmailAddressDetails} object for
         * each recipient listed on the **Cc** line of the message. Collection size limits:
         *
         * - Windows: 500 members
         *
         * - Classic Mac UI: 100 members
         *
         * - Web browser: 20 members
         *
         * - New Mac UI, Android: No limit
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        cc: EmailAddressDetails[];
        /**
         * Gets an identifier for the email conversation that contains a particular message.
         *
         * You can get an integer for this property if your mail app is activated in read forms or responses in compose forms.
         * If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change
         * and that value you obtained earlier will no longer apply.
         *
         * You get null for this property for a new item in a compose form.
         * If the user sets a subject and saves the item, the `conversationId` property will return a value.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        conversationId: string;
        /**
         * Gets the date and time that an item was created.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        dateTimeCreated: Date;
        /**
         * Gets the date and time that an item was last modified.
         *
         * **Note**: This member is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        dateTimeModified: Date;
        /**
         * Gets the date and time that the appointment is to end.
         *
         * The `end` property is a `Date` object expressed as a Coordinated Universal Time (UTC) date and time value.
         * You can use the `convertToLocalClientTime` method to convert the `end` property value to the client's local date and time.
         *
         * When you use the `Time.setAsync` method to set the end time, you should use the `convertToUtcClientTime` method to convert the local time on
         * the client to UTC for the server.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        end: Date;
        /**
         * Gets the email address of the sender of a message.
         *
         * The `from` and `sender` properties represent the same person unless the message is sent by a delegate.
         * In that case, the `from` property represents the delegator, and the `sender` property represents the delegate.
         *
         * **Note**: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is undefined.
         *
         * The `from` property returns an `EmailAddressDetails` object.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        from: EmailAddressDetails;
        /**
         * Gets the internet message identifier for an email message.
         *
         * **Important**: In the **Sent Items** folder, the `internetMessageId` may not be available yet on recently sent items. In that case,
         * consider using {@link https://learn.microsoft.com/office/dev/add-ins/outlook/web-services | Exchange Web Services} to get this
         * {@link https://learn.microsoft.com/exchange/client-developer/web-service-reference/internetmessageid | property from the server}.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        internetMessageId: string;
        /**
         * Gets the Exchange Web Services item class of the selected item.
         *
         * You can create custom message classes that extends a default message class, for example, a custom appointment message class
         * `IPM.Appointment.Contoso`.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * The `itemClass` property specifies the message class of the selected item.
         * The following are the default message classes for the message or appointment item.
         *
         * <table>
         *   <tr>
         *     <th>Type</th>
         *     <th>Description</th>
         *     <th>Item Class</th>
         *   </tr>
         *   <tr>
         *     <td>Appointment items</td>
         *     <td>These are calendar items of the item class IPM.Appointment or IPM.Appointment.Occurrence.</td>
         *     <td>IPM.Appointment, IPM.Appointment.Occurrence</td>
         *   </tr>
         *   <tr>
         *     <td>Message items</td>
         *     <td>These include email messages that have the default message class IPM.Note, and meeting requests, responses, and cancellations, that use IPM.Schedule.Meeting as the base message class.</td>
         *     <td>IPM.Note, IPM.Schedule.Meeting.Request, IPM.Schedule.Meeting.Neg, IPM.Schedule.Meeting.Pos, IPM.Schedule.Meeting.Tent, IPM.Schedule.Meeting.Canceled</td>
         *   </tr>
         * </table>
         *
         */
        itemClass: string;
        /**
         * Gets the {@link https://learn.microsoft.com/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange | Exchange Web Services item identifier}
         * for the current item.
         *
         * The `itemId` property is not available in compose mode.
         * If an item identifier is required, the `saveAsync` method can be used to save the item to the store, which will return the item identifier
         * in the `asyncResult.value` parameter in the callback function.
         *
         * **Note**: The identifier returned by the `itemId` property is the same as the
         * {@link https://learn.microsoft.com/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange | Exchange Web Services item identifier}.
         * The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.
         * Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`.
         * For more details, see {@link https://learn.microsoft.com/office/dev/add-ins/outlook/use-rest-api#get-the-item-id | Use the Outlook REST APIs from an Outlook add-in}.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        itemId: string;
        /**
         * Gets the type of item that an instance represents.
         *
         * The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the item object instance is a message or
         * an appointment.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        itemType: MailboxEnums.ItemType | string;
        /**
         * Gets the location of a meeting request.
         *
         * The `location` property returns a string that contains the location of the appointment.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        location: string;
        /**
         * Gets the subject of an item, with all prefixes removed (including RE: and FWD:).
         *
         * The `normalizedSubject` property gets the subject of the item, with any standard prefixes (such as RE: and FW:) that are added by
         * email programs. To get the subject of the item with the prefixes intact, use the `subject` property.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        normalizedSubject: string;
        /**
         * Gets the notification messages for an item.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        notificationMessages: NotificationMessages;
        /**
         * Gets the recurrence pattern of an appointment. Gets the recurrence pattern of a meeting request.
         * Read and compose modes for appointment items. Read mode for meeting request items.
         *
         * The `recurrence` property returns a `Recurrence` object for recurring appointments or meetings requests if an item is a series or an instance
         * in a series. `null` is returned for single appointments and meeting requests of single appointments.
         * `undefined` is returned for messages that are not meeting requests.
         *
         * **Note**: Meeting requests have an itemClass value of `IPM.Schedule.Meeting.Request`.
         *
         * **Note**: If the `recurrence` object is null, this indicates that the object is a single appointment or a meeting request of a single appointment
         * and NOT a part of a series.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        recurrence: Recurrence;
        /**
         * Gets the ID of the series that an instance belongs to.
         *
         * In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item
         * that this item belongs to. However, on iOS and Android, the `seriesId` returns the REST ID of the parent item.
         *
         * **Note**: The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.
         * The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.
         * Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`.
         * For more details, see {@link https://learn.microsoft.com/office/dev/add-ins/outlook/use-rest-api | Use the Outlook REST APIs from an Outlook add-in}.
         *
         * The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests
         * and returns `undefined` for any other items that are not meeting requests.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        seriesId: string;
        /**
         * Gets the email address of the sender of an email message.
         *
         * The `from` and `sender` properties represent the same person unless the message is sent by a delegate.
         * In that case, the `from` property represents the delegator, and the `sender` property represents the delegate.
         *
         * **Note**: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is undefined.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        sender: EmailAddressDetails;
        /**
         * Gets the date and time that the appointment is to begin.
         *
         * The `start` property is a `Date` object expressed as a Coordinated Universal Time (UTC) date and time value.
         * You can use the `convertToLocalClientTime` method to convert the value to the client's local date and time.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        start: Date;
        /**
         * Gets the description that appears in the subject field of an item.
         *
         * The `subject` property gets or sets the entire subject of the item, as sent by the email server.
         *
         * The `subject` property returns a string. Use the `normalizedSubject` property to get the subject minus any leading prefixes such as RE: and FW:.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        subject: string;
        /**
         * Provides access to the recipients on the **To** line of a message. The type of object and level of access depend on the mode of the
         * current item.
         *
         * The `to` property returns an array that contains an {@link Office.EmailAddressDetails | EmailAddressDetails} object for
         * each recipient listed on the **To** line of the message. Collection size limits:
         *
         * - Windows: 500 members
         *
         * - Classic Mac UI: 100 members
         *
         * - Web browser: 20 members
         *
         * - New Mac UI, Android: No limit
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        to: EmailAddressDetails[];

        /**
         * Adds an event handler for a supported event. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Item object model
         * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param eventType - The event that should invoke the handler.
         * @param handler - The function to handle the event. The function must accept a single parameter, which is an object literal.
         *                The `type` property on the parameter will match the eventType `parameter` passed to `addHandlerAsync`.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        addHandlerAsync(eventType: CommonAPI.EventType | string, handler: any, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Adds an event handler for a supported event. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Item object model
         * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param eventType - The event that should invoke the handler.
         * @param handler - The function to handle the event. The function must accept a single parameter, which is an object literal.
         *                The `type` property on the parameter will match the eventType `parameter` passed to `addHandlerAsync`.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        addHandlerAsync(eventType: CommonAPI.EventType | string, handler: any, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Displays a reply form that includes either the sender and all recipients of the selected message or the organizer and all attendees of the
         * selected appointment.
         *
         * In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2-column or 1-column view.
         *
         * If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.
         *
         * When attachments are specified in the `formData.attachments` parameter, Outlook attempts to download all attachments and attach them to the
         * reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *                   OR a {@link Office.ReplyFormData | ReplyFormData} object that contains body or attachment data and a callback function.
         */
        displayReplyAllForm(formData: string | ReplyFormData): void;
        /**
         * Displays a reply form that includes either the sender and all recipients of the selected message or the organizer and all attendees of the
         * selected appointment.
         *
         * In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2-column or 1-column view.
         *
         * If any of the string parameters exceed their limits, `displayReplyAllFormAsync` throws an exception.
         *
         * When attachments are specified in the `formData.attachments` parameter, Outlook attempts to download all attachments and attach them to the
         * reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.9]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *                   OR a {@link Office.ReplyFormData | ReplyFormData} object that contains body or attachment data and a callback function.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        displayReplyAllFormAsync(formData: string | ReplyFormData, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Displays a reply form that includes either the sender and all recipients of the selected message or the organizer and all attendees of the
         * selected appointment.
         *
         * In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2-column or 1-column view.
         *
         * If any of the string parameters exceed their limits, `displayReplyAllFormAsync` throws an exception.
         *
         * When attachments are specified in the `formData.attachments` parameter, Outlook attempts to download all attachments and attach them to the
         * reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.9]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *                   OR a {@link Office.ReplyFormData | ReplyFormData} object that contains body or attachment data and a callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        displayReplyAllFormAsync(formData: string | ReplyFormData, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.
         *
         * In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2-column or 1-column view.
         *
         * If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.
         *
         * When attachments are specified in the `formData.attachments` parameter, Outlook attempts to download all attachments and attach them to the
         * reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *                   OR a {@link Office.ReplyFormData | ReplyFormData} object that contains body or attachment data and a callback function.
         */
        displayReplyForm(formData: string | ReplyFormData): void;
        /**
         * Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.
         *
         * In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2-column or 1-column view.
         *
         * If any of the string parameters exceed their limits, `displayReplyFormAsync` throws an exception.
         *
         * When attachments are specified in the `formData.attachments` parameter, Outlook attempts to download all attachments and attach them to the
         * reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.9]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *                   OR a {@link Office.ReplyFormData | ReplyFormData} object that contains body or attachment data and a callback function.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        displayReplyFormAsync(formData: string | ReplyFormData, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.
         *
         * In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2-column or 1-column view.
         *
         * If any of the string parameters exceed their limits, `displayReplyFormAsync` throws an exception.
         *
         * When attachments are specified in the `formData.attachments` parameter, Outlook attempts to download all attachments and attach them to the
         * reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.9]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param formData - A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB
         *                   OR a {@link Office.ReplyFormData | ReplyFormData} object that contains body or attachment data and a callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        displayReplyFormAsync(formData: string | ReplyFormData, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Gets all the internet headers for the message as a string.
         *
         * To learn more, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/internet-headers | Get and set internet headers on a message in an Outlook add-in}.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         *                On success, the internet headers data is provided in the `asyncResult.value` property as a string.
         *                Refer to {@link https://tools.ietf.org/html/rfc2183 | RFC 2183} for the formatting information of the returned string value.
         *                If the call fails, the `asyncResult.error` property will contain an error code with the reason for the failure.
         */
        getAllInternetHeadersAsync(options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Gets all the internet headers for the message as a string.
         *
         * To learn more, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/internet-headers | Get and set internet headers on a message in an Outlook add-in}.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         *                On success, the internet headers data is provided in the `asyncResult.value` property as a string.
         *                Refer to {@link https://tools.ietf.org/html/rfc2183 | RFC 2183} for the formatting information of the returned string value.
         *                If the call fails, the `asyncResult.error` property will contain an error code with the reason for the failure.
         */
        getAllInternetHeadersAsync(callback?: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Gets an attachment from a message or appointment and returns it as an `AttachmentContent` object.
         *
         * The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item. As a best practice, you should get
         * the attachment's identifier from an {@link Office.MessageRead.attachments | item.attachments} call, then in the same session, use that identifier
         * to retrieve the attachment. In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.
         * A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to
         * continue in a separate window.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * **Errors**:
         *
         * - `AttachmentTypeNotSupported`: The attachment type isn't supported. Unsupported types include embedded images in Rich Text Format,
         *                               or item attachment types other than email or calendar items (such as a contact or task item).
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param attachmentId - The identifier of the attachment you want to get.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object. If the call fails, the `asyncResult.error` property will contain
         *                an error code with the reason for the failure.
         */
        getAttachmentContentAsync(attachmentId: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<AttachmentContent>) => void): void;
        /**
         * Gets an attachment from a message or appointment and returns it as an `AttachmentContent` object.
         *
         * The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item. As a best practice, you should get
         * the attachment's identifier from an {@link Office.MessageRead.attachments | item.attachments} call, then in the same session, use that identifier
         * to retrieve the attachment. In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.
         * A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to
         * continue in a separate window.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * **Errors**:
         *
         * - `AttachmentTypeNotSupported`: The attachment type isn't supported. Unsupported types include embedded images in Rich Text Format,
         *                               or item attachment types other than email or calendar items (such as a contact or task item).
         *
         * - `InvalidAttachmentId`: The attachment identifier does not exist.
         *
         * @param attachmentId - The identifier of the attachment you want to get.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object. If the call fails, the `asyncResult.error` property will contain
         *                an error code with the reason for the failure.
         */
        getAttachmentContentAsync(attachmentId: string, callback?: (asyncResult: CommonAPI.AsyncResult<AttachmentContent>) => void): void;
        /**
         * Gets the entities found in the selected item's body.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        getEntities(): Entities;
        /**
         * Gets an array of all the entities of the specified entity type found in the selected item's body.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @returns
         * If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns `null`.
         * If no entities of the specified type are present in the item's body, the method returns an empty array.
         * Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param entityType - One of the `EntityType` enumeration values.
         *
         * While the minimum permission level to use this method is **restricted**, some entity types require **read item** to access, as specified in the
         * following table.
         *
         * <table>
         *   <tr>
         *     <th>Value of entityType</th>
         *     <th>Type of objects in returned array</th>
         *     <th>Required Permission Level</th>
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
        getEntitiesByType(entityType: MailboxEnums.EntityType | string): (string | Contact | MeetingSuggestion | PhoneNumber | TaskSuggestion)[];
        /**
         * Returns well-known entities in the selected item that pass the named filter defined in an XML manifest file.
         * 
         * **Note**: This method is used with the {@link https://learn.microsoft.com/office/dev/add-ins/outlook/activation-rules | activation rules feature for Outlook add-ins}, which isn't supported by the {@link https://learn.microsoft.com/office/dev/add-ins/develop/json-manifest-overview | Teams manifest for Office Add-ins (preview)}.
         *
         * The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the `ItemHasKnownEntity` rule element
         * in the manifest XML file with the specified `FilterName` element value.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @returns If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter,
         * the method returns `null`.
         * If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match,
         * the method return an empty array.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param name - The name of the `ItemHasKnownEntity` rule element that defines the filter to match.
         */
        getFilteredEntitiesByName(name: string): (string | Contact | MeetingSuggestion | PhoneNumber | TaskSuggestion)[];
        /**
         * Gets initialization data passed when the add-in is {@link https://learn.microsoft.com/outlook/actionable-messages/invoke-add-in | activated by an actionable message}.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         *                 On success, the initialization context data is provided as a string (or an empty string if there's no initialization context)
         *                 in the `asyncResult.value` property.
         */
        getInitializationContextAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Gets initialization data passed when the add-in is {@link https://learn.microsoft.com/outlook/actionable-messages/invoke-add-in | activated by an actionable message}.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         *                 On success, the initialization context data is provided as a string (or an empty string if there's no initialization context)
         *                 in the `asyncResult.value` property.
         */
        getInitializationContextAsync(callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Returns string values in the selected item that match the regular expressions defined in an XML manifest file.
         * 
         * **Note**: This method is used with the {@link https://learn.microsoft.com/office/dev/add-ins/outlook/activation-rules | activation rules feature for Outlook add-ins}, which isn't supported by the {@link https://learn.microsoft.com/office/dev/add-ins/develop/json-manifest-overview | Teams manifest for Office Add-ins (preview)}.
         *
         * The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or
         * `ItemHasKnownEntity` rule element in the manifest XML file.
         * For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule.
         * The `PropertyName` simple type defines the supported properties.
         *
         * If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter
         * the body and should not attempt to return the entire body of the item.
         * Using a regular expression such as .* to obtain the entire body of an item does not always return the expected results.
         * Instead, use the `Body.getAsync` method to retrieve the entire body.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @returns
         * An object that contains arrays of strings that match the regular expressions defined in the manifest XML file.
         * The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule
         * or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        getRegExMatches(): any;
        /**
         * Returns string values in the selected item that match the named regular expression defined in an XML manifest file.
         * 
         * **Note**: This method is used with the {@link https://learn.microsoft.com/office/dev/add-ins/outlook/activation-rules | activation rules feature for Outlook add-ins}, which isn't supported by the {@link https://learn.microsoft.com/office/dev/add-ins/develop/json-manifest-overview | Teams manifest for Office Add-ins (preview)}.
         *
         * The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the
         * `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.
         *
         * If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter
         * the body and should not attempt to return the entire body of the item.
         * Using a regular expression such as .* to obtain the entire body of an item does not always return the expected results.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @returns
         * An array that contains the strings that match the regular expression defined in the manifest XML file.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param name - The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.
         */
        getRegExMatchesByName(name: string): string[];
        /**
         * Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to contextual add-ins.
         * 
         * **Note**: This method is used with the {@link https://learn.microsoft.com/office/dev/add-ins/outlook/activation-rules | activation rules feature for Outlook add-ins}, which isn't supported by the {@link https://learn.microsoft.com/office/dev/add-ins/develop/json-manifest-overview | Teams manifest for Office Add-ins (preview)}.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.6]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param name - The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.
         */
        getSelectedEntities(): Entities;
        /**
         * Returns string values in a highlighted match that match the regular expressions defined in an XML manifest file.
         * Highlighted matches apply to contextual add-ins.
         * 
         * **Note**: This method is used with the {@link https://learn.microsoft.com/office/dev/add-ins/outlook/activation-rules | activation rules feature for Outlook add-ins}, which isn't supported by the {@link https://learn.microsoft.com/office/dev/add-ins/develop/json-manifest-overview | Teams manifest for Office Add-ins (preview)}.
         *
         * The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in
         * each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file.
         * For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule.
         * The `PropertyName` simple type defines the supported properties.
         *
         * If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body
         * and should not attempt to return the entire body of the item.
         * Using a regular expression such as .* to obtain the entire body of an item does not always return the expected results.
         * Instead, use the `Body.getAsync` method to retrieve the entire body.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @returns
         * An object that contains arrays of strings that match the regular expressions defined in the manifest XML file.
         * The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or
         * the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.
         *
         * @remarks
         * [Api set: Mailbox 1.6]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         */
        getSelectedRegExMatches(): any;
        /**
         * Gets the properties of an appointment or message in a shared folder or shared mailbox (now in preview).
         *
         * For more information around using this API, see the
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/delegate-access | shared folders and shared mailbox} article.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. The `value` property of the result is the properties of the shared item.
         */
        getSharedPropertiesAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<SharedProperties>) => void): void;
        /**
         * Gets the properties of an appointment or message in a shared folder or shared mailbox (now in preview).
         *
         * For more information around using this API, see the
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/delegate-access | shared folders and shared mailbox} article.
         *
         * **Note**: This method is not supported in Outlook on iOS or Android.
         *
         * @remarks
         * [Api set: Mailbox 1.8]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. The `value` property of the result is the properties of the shared item.
         */
        getSharedPropertiesAsync(callback: (asyncResult: CommonAPI.AsyncResult<SharedProperties>) => void): void;
        /**
         * Asynchronously loads custom properties for this add-in on the selected item.
         *
         * Custom properties are stored as key-value pairs on a per-app, per-item basis.
         * This method returns a {@link Office.CustomProperties | CustomProperties} object in the callback, which provides methods to access the custom properties specific to the
         * current item and the current add-in. Custom properties aren't encrypted on the item, so this shouldn't be used as secure storage.
         *
         * The custom properties are provided as a `CustomProperties` object in the `asyncResult.value` property.
         * This object can be used to get, set, save, and remove custom properties from the mail item.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * To learn more about custom properties, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/metadata-for-an-outlook-add-in | Get and set add-in metadata for an Outlook add-in}.
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         * @param userContext - Optional. Developers can provide any object they wish to access in the callback function.
         *                    This object can be accessed by the `asyncResult.asyncContext` property in the callback function.
         */
        loadCustomPropertiesAsync(callback: (asyncResult: CommonAPI.AsyncResult<CustomProperties>) => void, userContext?: any): void;
        /**
         * Removes the event handlers for a supported event type. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Item object model
         * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param eventType - The event that should revoke the handler.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        removeHandlerAsync(eventType: CommonAPI.EventType | string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Removes the event handlers for a supported event type. **Note**: Events are only available with task pane implementation.
         *
         * For supported events, refer to the Item object model
         * {@link https://learn.microsoft.com/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#events | events section}.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Message Read
         *
         * @param eventType - The event that should revoke the handler.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        removeHandlerAsync(eventType: CommonAPI.EventType | string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * Provides methods to get and set the all-day event status of a meeting in an Outlook add-in.
     *
     * @remarks
     * [Api set: Mailbox preview]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
     *
     * @beta
     */
    export interface IsAllDayEvent {
        /**
         * Gets the boolean value indicating whether the event is all day or not.
         *
         * @remarks
         * [Api set: Mailbox preview]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         *
         * @beta
         */
        getAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<boolean>) => void): void;
        /**
         * Gets the boolean value indicating whether the event is all day or not.
         *
         * @remarks
         * [Api set: Mailbox preview]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         *
         * @beta
         */
        getAsync(callback: (asyncResult: CommonAPI.AsyncResult<boolean>) => void): void;
        /**
         * Sets the all-day event status of an appointment.
         *
         * @remarks
         * [Api set: Mailbox preview]
         * 
         * If an appointment is marked as an all-day event:
         * - Start and end time will be marked as 12:00 AM (just like in the Outlook UI). Start time will return 12:00 AM and end time will be 12:00 AM the next day.
         *
         * **{@link  https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param isAllDayEvent - boolean value to set the all day event status.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         *
         * @beta
         */
        setAsync(isAllDayEvent: boolean, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Sets the all-day event status of an appointment.
         *
         * @remarks
         * [Api set: Mailbox preview]
         *
         * If an appointment is marked as an all-day event:
         * - Start and end time will be marked as 12:00 AM (just like in the Outlook UI). Start time will return 12:00 AM and end time will be 12:00 AM the next day.
         *
         * **{@link  https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param isAllDayEvent - boolean value to set the all day event status.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         *
         * @beta
         */
        setAsync(isAllDayEvent: boolean, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * The definition of the action for a notification message.
     *
     * **Important**: In modern Outlook on the web, the `NotificationMessageAction` object is available in Compose mode only.
     *
     * @remarks
     * [Api set: Mailbox 1.10]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface NotificationMessageAction {
        /**
         * The type of action to be performed.
         * `ActionType.ShowTaskPane` is the only supported action.
         */
        actionType: string | MailboxEnums.ActionType;
        /**
         * The text of the action link.
         */
        actionText: string;
        /**
         * The button defined in the manifest.
         */
        commandId: string;
        /**
         * Any JSON data the action button needs to pass on.
         * This data can be retrieved by calling `item.getInitializationContextAsync`.
         */
        contextData: any;
    }
    /**
     * An array of `NotificationMessageDetails` objects are returned by the `NotificationMessages.getAllAsync` method.
     *
     * @remarks
     * [Api set: Mailbox 1.3]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface NotificationMessageDetails {
        /**
         * The identifier for the notification message.
         */
        key?: string;
        /**
         * Specifies the `ItemNotificationMessageType` of message.
         *
         * If type is `ProgressIndicator` or `ErrorMessage`, an icon is automatically supplied
         * and the message is not persistent. Therefore the icon and persistent properties are not valid for these types of messages.
         * Including them will result in an `ArgumentException`.
         *
         * If type is `ProgressIndicator`, the developer should remove or replace the progress indicator when the action is complete.
         */
        type: MailboxEnums.ItemNotificationMessageType | string;
        /**
         * A reference to an icon that is defined in the manifest. It appears in the infobar area.
         * It is applicable if the type is `InformationalMessage`, and is required if the type is `InsightMessage`.
         * Specifying this parameter for an unsupported type results in an exception.
         *
         * **Note**: At present, the custom icon is displayed in Outlook on Windows only and not on other clients (e.g., Mac, web browser).
         */
        icon?: string;
        /**
         * The text of the notification message. Maximum length is 150 characters.
         * If the developer passes in a longer string, an `ArgumentOutOfRange` exception is thrown.
         */
        message: string;
        /**
         * Specifies if the message should be persistent. Only applicable when type is `InformationalMessage`.
         * If true, the message remains until removed by this add-in or dismissed by the user.
         * If false, it is removed when the user navigates to a different item.
         * For error notifications, the message persists until the user sees it once.
         * Specifying this parameter for an unsupported type throws an exception.
         */
        persistent?: Boolean;
        /**
         * Specifies actions for the message. Limit: 1 action. This limit doesn't count the "Dismiss" action which is included by default.
         * Only applicable when the type is `InsightMessage`.
         * Specifying this property for an unsupported type or including too many actions throws an error.
         *
         * **Important**: In modern Outlook on the web, the `actions` property is available in Compose mode only.
         *
         * @remarks
         * [Api set: Mailbox 1.10]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        actions?: NotificationMessageAction[];
    }
    /**
     * The `NotificationMessages` object is returned as the `notificationMessages` property of an item.
     *
     * @remarks
     * [Api set: Mailbox 1.3]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface NotificationMessages {
        /**
         * Adds a notification to an item.
         *
         * There are a maximum of 5 notifications per message. Setting more will return a `NumberOfNotificationMessagesExceeded` error.
         *
         * **Important**:
         *
         * - Only one notification of type {@link https://learn.microsoft.com/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype#fields | InsightMessage}
         * is allowed per add-in. Attempting to add more will throw an error.
         *
         * - In modern Outlook on the web, you can add an `InsightMessage` notification only in Compose mode.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param key - A developer-specified key used to reference this notification message.
         *             Developers can use it to modify this message later. It can't be longer than 32 characters.
         * @param JSONmessage - A JSON object that contains the notification message to be added to the item.
         *                    It contains a `NotificationMessageDetails` object.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         */
        addAsync(key: string, JSONmessage: NotificationMessageDetails, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Adds a notification to an item.
         *
         * There are a maximum of 5 notifications per message. Setting more will return a `NumberOfNotificationMessagesExceeded` error.
         *
         * **Important**:
         *
         * - Only one notification of type {@link https://learn.microsoft.com/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype#fields | InsightMessage}
         * is allowed per add-in. Attempting to add more will throw an error.
         *
         * - In modern Outlook on the web, you can add an `InsightMessage` notification only in Compose mode.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param key - A developer-specified key used to reference this notification message.
         *             Developers can use it to modify this message later. It can't be longer than 32 characters.
         * @param JSONmessage - A JSON object that contains the notification message to be added to the item.
         *                    It contains a `NotificationMessageDetails` object.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         */
        addAsync(key: string, JSONmessage: NotificationMessageDetails, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Returns all keys and messages for an item.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. The `value` property of the result is an array of `NotificationMessageDetails` objects.
         */
        getAllAsync(options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<NotificationMessageDetails[]>) => void): void;
        /**
         * Returns all keys and messages for an item.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. The `value` property of the result is an array of `NotificationMessageDetails` objects.
         */
        getAllAsync(callback?: (asyncResult: CommonAPI.AsyncResult<NotificationMessageDetails[]>) => void): void;
        /**
         * Removes a notification message for an item.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param key - The key for the notification message to remove.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         */
        removeAsync(key: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Removes a notification message for an item.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param key - The key for the notification message to remove.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         */
        removeAsync(key: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Replaces a notification message that has a given key with another message.
         *
         * If a notification message with the specified key doesn't exist, `replaceAsync` will add the notification.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param key - The key for the notification message to replace. It can't be longer than 32 characters.
         * @param JSONmessage - A JSON object that contains the new notification message to replace the existing message.
         *                    It contains a `NotificationMessageDetails` object.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         */
        replaceAsync(key: string, JSONmessage: NotificationMessageDetails, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Replaces a notification message that has a given key with another message.
         *
         * If a notification message with the specified key doesn't exist, `replaceAsync` will add the notification.
         *
         * @remarks
         * [Api set: Mailbox 1.3]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param key - The key for the notification message to replace. It can't be longer than 32 characters.
         * @param JSONmessage - A JSON object that contains the new notification message to replace the existing message.
         *                    It contains a `NotificationMessageDetails` object.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         */
        replaceAsync(key: string, JSONmessage: NotificationMessageDetails, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * Provides the updated Office theme that raised the `Office.EventType.OfficeThemeChanged` event.
     *
     * @remarks
     * [Api set: Mailbox preview]
     *
     * @beta
     */
    export interface OfficeThemeChangedEventArgs {
        /**
         * Gets the updated Office theme.
         *
         * @remarks
         * [Api set: Mailbox preview]
         *
         * @beta
         */
        officeTheme: CommonAPI.OfficeTheme;
        /**
         * Gets the type of the event. For details, refer to {@link https://learn.microsoft.com/javascript/api/office/office.eventtype | Office.EventType}.
         *
         * @remarks
         * [Api set: Mailbox preview]
         *
         * @beta
         */
        type: "officeThemeChanged";
    }
    /**
     * Represents the appointment organizer, even if an alias or a delegate was used to create the appointment.
     * This object provides a method to get the organizer value of an appointment in an Outlook add-in.
     *
     * @remarks
     * [Api set: Mailbox 1.7]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
     */
    export interface Organizer {
        /**
         * Gets the organizer value of an appointment as an {@link Office.EmailAddressDetails | EmailAddressDetails} object
         * in the `asyncResult.value` property.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Important**: A `recipientType` property value isn't returned by the getAsync method.
         * The appointment organizer is always a user whose email address is on the Exchange server.
         * 
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                  `asyncResult`, which is an `AsyncResult` object. The `value` property of the result is the appointment's organizer value,
         *                  as an `EmailAddressDetails` object.
         */
        getAsync(options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<EmailAddressDetails>) => void): void;
        /**
         * Gets the organizer value of an appointment as an {@link Office.EmailAddressDetails | EmailAddressDetails} object
         * in the `asyncResult.value` property.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * **Important**: A `recipientType` property value isn't returned by the getAsync method.
         * The appointment organizer is always a user whose email address is on the Exchange server.
         *
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                  `asyncResult`, which is an `AsyncResult` object. The `value` property of the result is the appointment's organizer value,
         *                  as an `EmailAddressDetails` object.
         */
        getAsync(callback?: (asyncResult: CommonAPI.AsyncResult<EmailAddressDetails>) => void): void;
    }
    /**
     * Represents a phone number identified in an item. Read mode only.
     *
     * An array of `PhoneNumber` objects containing the phone numbers found in an email message is returned in the `phoneNumbers` property of the
     * `Entities` object that is returned when you call the `getEntities` method on the selected item.
     *
     * @remarks
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
     */
    export interface PhoneNumber {
        /**
         * Gets a string containing a phone number. This string contains only the digits of the telephone number and excludes characters
         * like parentheses and hyphens, if they exist in the original item.
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
     * Represents recipients of an item. Compose mode only.
     *
     * @remarks
     * [Api set: Mailbox 1.1]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
     */
    export interface Recipients {
        /**
         * Adds a recipient list to the existing recipients for an appointment or message.
         *
         * The recipients parameter can be an array of one of the following:
         *
         * - Strings containing SMTP email addresses
         *
         * - {@link Office.EmailUser | EmailUser} objects
         *
         * - {@link Office.EmailAddressDetails | EmailAddressDetails} objects
         *
         * Maximum number that can be added:
         *
         * - Windows: 100 recipients.
         * **Note**: Can call API repeatedly but the maximum number of recipients in the target field on the item is 500 recipients.
         *
         * - Classic Mac UI, web browser: 100 recipients
         *
         * - New Mac UI, Android: No limit
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `NumberOfRecipientsExceeded`: The number of recipients exceeded 100 entries.
         *
         * @param recipients - The recipients to add to the recipients list.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. If adding the recipients fails, the `asyncResult.error` property will contain an error code.
         */
        addAsync(recipients: (string | EmailUser | EmailAddressDetails)[], options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Adds a recipient list to the existing recipients for an appointment or message.
         *
         * The recipients parameter can be an array of one of the following:
         *
         * - Strings containing SMTP email addresses
         *
         * - {@link Office.EmailUser | EmailUser} objects
         *
         * - {@link Office.EmailAddressDetails | EmailAddressDetails} objects
         *
         * Maximum number that can be added:
         *
         * - Windows: 100 recipients.
         * **Note**: Can call API repeatedly but the maximum number of recipients in the target field on the item is 500 recipients.
         *
         * - Classic Mac UI, web browser: 100 recipients
         *
         * - New Mac UI, Android: No limit
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `NumberOfRecipientsExceeded`: The number of recipients exceeded 100 entries.
         *
         * @param recipients - The recipients to add to the recipients list.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. If adding the recipients fails, the `asyncResult.error` property will contain an error code.
         */
        addAsync(recipients: (string | EmailUser | EmailAddressDetails)[], callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Gets a recipient list for an appointment or message.
         *
         * When the call completes, the `asyncResult.value` property will contain an array of {@link Office.EmailAddressDetails | EmailAddressDetails}
         * objects. Collection size limits:
         *
         * - Windows, classic Mac UI, web browser: 500 members
         *
         * - New Mac UI, Android: No limit
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * **Important**: In Outlook on the web and on Windows, if a user creates a new message by activating a contact's email address link from their contact
         * or profile card, your add-in's `Recipients.getAsync` call returns the contact's email address in the `displayName` property of the associated
         * `EmailAddressDetails` object instead of the contact's saved name. For more details, refer to the
         * {@link https://github.com/OfficeDev/office-js/issues/2201 | related GitHub issue}.
         * 
         * **Important**: The `getAsync` method only returns recipients resolved by the Outlook client. A resolved recipient has the following characteristics.
         * 
         * - If the recipient has a saved entry in the sender's address book, Outlook resolves the email address to the recipient's saved display name.
         * 
         * - A Teams meeting status icon appears before the recipient's name or email address.
         * 
         * - A semicolon appears after the recipient's name or email address.
         * 
         * - The recipient's name or email address is underlined or enclosed in a box.
         * 
         * To resolve an email address once it's added to a mail item, the sender must use the **Tab** key or select a suggested contact or email address from 
         * the auto-complete list.
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. The `value` property of the result is an array of `EmailAddressDetails` objects.
         */
        getAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<EmailAddressDetails[]>) => void): void;
        /**
         * Gets a recipient list for an appointment or message.
         *
         * When the call completes, the `asyncResult.value` property will contain an array of {@link Office.EmailAddressDetails | EmailAddressDetails}
         * objects. Collection size limits:
         *
         * - Windows, classic Mac UI, web browser: 500 members
         *
         * - New Mac UI, Android: No limit
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Important**: In Outlook on the web and on Windows, if a user creates a new message by activating a contact's email address link from their contact
         * or profile card, your add-in's `Recipients.getAsync` call returns the contact's email address in the `displayName` property of the associated
         * `EmailAddressDetails` object instead of the contact's saved name. For more details, refer to the
         * {@link https://github.com/OfficeDev/office-js/issues/2201 | related GitHub issue}.
         * 
         * **Important**: The `getAsync` method only returns recipients resolved by the Outlook client. A resolved recipient has the following characteristics.
         * 
         * - If the recipient has a saved entry in the sender's address book, Outlook resolves the email address to the recipient's saved display name.
         * 
         * - A Teams meeting status icon appears before the recipient's name or email address.
         * 
         * - A semicolon appears after the recipient's name or email address.
         * 
         * - The recipient's name or email address is underlined or enclosed in a box.
         * 
         * To resolve an email address once it's added to a mail item, the sender must use the **Tab** key or select a suggested contact or email address from 
         * the auto-complete list.
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. The `value` property of the result is an array of `EmailAddressDetails` objects.
         */
        getAsync(callback: (asyncResult: CommonAPI.AsyncResult<EmailAddressDetails[]>) => void): void;
        /**
         * Sets a recipient list for an appointment or message.
         *
         * The `setAsync` method overwrites the current recipient list.
         *
         * The recipients parameter can be an array of one of the following:
         *
         * - Strings containing SMTP email addresses
         *
         * - {@link Office.EmailUser | EmailUser} objects
         *
         * - {@link Office.EmailAddressDetails | EmailAddressDetails} objects
         *
         * Maximum number that can be set:
         *
         * - Windows: 100 recipients.
         * **Note**: Can call API repeatedly but the maximum number of recipients in the target field on the item is 500 recipients.
         *
         * - Classic Mac UI, web browser: 100 recipients
         *
         * - New Mac UI, Android: No limit
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `NumberOfRecipientsExceeded`: The number of recipients exceeded 100 entries.
         *
         * @param recipients - The recipients to add to the recipients list.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If setting the recipients fails the `asyncResult.error` property will contain a code that
         *                 indicates any error that occurred while adding the data.
         */
        setAsync(recipients: (string | EmailUser | EmailAddressDetails)[], options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Sets a recipient list for an appointment or message.
         *
         * The `setAsync` method overwrites the current recipient list.
         *
         * The recipients parameter can be an array of one of the following:
         *
         * - Strings containing SMTP email addresses
         *
         * - {@link Office.EmailUser | EmailUser} objects
         *
         * - {@link Office.EmailAddressDetails | EmailAddressDetails} objects
         *
         * Maximum number that can be set:
         *
         * - Windows: 100 recipients.
         * **Note**: Can call API repeatedly but the maximum number of recipients in the target field on the item is 500 recipients.
         *
         * - Classic Mac UI, web browser: 100 recipients
         *
         * - New Mac UI, Android: No limit
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `NumberOfRecipientsExceeded`: The number of recipients exceeded 100 entries.
         *
         * @param recipients - The recipients to add to the recipients list.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`. If setting the recipients fails the `asyncResult.error` property will contain a code that
         *                 indicates any error that occurred while adding the data.
         */
        setAsync(recipients: (string | EmailUser | EmailAddressDetails)[], callback: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * Provides change status of recipients fields when the `Office.EventType.RecipientsChanged` event is raised.
     *
     * @remarks
     * [Api set: Mailbox 1.7]
     */
    export interface RecipientsChangedEventArgs {
        /**
         * Gets an object that indicates change state of recipients fields.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         */
        changedRecipientFields: RecipientsChangedFields;
        /**
         * Gets the type of the event. For details, refer to {@link https://learn.microsoft.com/javascript/api/office/office.eventtype | Office.EventType}.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         */
        type: "olkRecipientsChanged";
    }
    /**
     * Represents `RecipientsChangedEventArgs.changedRecipientFields` object.
     *
     * @remarks
     * [Api set: Mailbox 1.7]
     */
    export interface RecipientsChangedFields {
        /**
         * Gets if recipients in the **bcc** field were changed.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         */
        bcc: boolean
        /**
         * Gets if recipients in the **cc** field were changed.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         */
        cc: boolean;
        /**
         * Gets if optional attendees were changed.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         */
        optionalAttendees: boolean;
        /**
         * Gets if required attendees were changed.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         */
        requiredAttendees: boolean;
        /**
         * Gets if resources were changed.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         */
        resources: boolean;
        /**
         * Gets if recipients in the **to** field were changed.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         */
        to: boolean;
    }
    /**
     * The `Recurrence` object provides methods to get and set the recurrence pattern of appointments but only get the recurrence pattern of
     * meeting requests.
     * It will have a dictionary with the following keys: `seriesTime`, `recurrenceType`, `recurrenceProperties`, and `recurrenceTimeZone` (optional).
     *
     * @remarks
     * [Api set: Mailbox 1.7]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     *
     * **States**
     *
     * <table>
     *   <tr>
     *     <th>State</th>
     *     <th>Editable?</th>
     *     <th>Viewable?</th>
     *   </tr>
     *   <tr>
     *     <td>Appointment Organizer - Compose Series</td>
     *     <td>Yes (setAsync)</td>
     *     <td>Yes (getAsync)</td>
     *   </tr>
     *   <tr>
     *     <td>Appointment Organizer - Compose Instance</td>
     *     <td>No (setAsync returns error)</td>
     *     <td>Yes (getAsync)</td>
     *   </tr>
     *   <tr>
     *     <td>Appointment Attendee - Read Series</td>
     *     <td>No (setAsync not available)</td>
     *     <td>Yes (item.recurrence)</td>
     *   </tr>
     *   <tr>
     *     <td>Appointment Attendee - Read Instance</td>
     *     <td>No (setAsync not available)</td>
     *     <td>Yes (item.recurrence)</td>
     *   </tr>
     *   <tr>
     *     <td>Meeting Request - Read Series</td>
     *     <td>No (setAsync not available)</td>
     *     <td>Yes (item.recurrence)</td>
     *   </tr>
     *   <tr>
     *     <td>Meeting Request - Read Instance</td>
     *     <td>No (setAsync not available)</td>
     *     <td>Yes (item.recurrence)</td>
     *   </tr>
     * </table>
     */
    export interface Recurrence {
        /**
         * Gets or sets the properties of the recurring appointment series.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        recurrenceProperties?: RecurrenceProperties;
        /**
         * Gets or sets the properties of the recurring appointment series.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        recurrenceTimeZone?: RecurrenceTimeZone;
        /**
         * Gets or sets the type of the recurring appointment series.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        recurrenceType: MailboxEnums.RecurrenceType | string;
        /**
         * The {@link Office.SeriesTime | SeriesTime} object enables you to manage the start and end dates of the recurring appointment series and
         * the usual start and end times of instances. **This object is not in UTC time.**
         * Instead, it is set in the time zone specified by the `recurrenceTimeZone` value or defaulted to the item's time zone.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        seriesTime: SeriesTime;

        /**
         * Returns the current recurrence object of an appointment series.
         *
         * This method returns the entire `Recurrence` object for the appointment series.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object. The `value` property of the result is a `Recurrence` object.
         */
        getAsync(options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<Recurrence>) => void): void;
        /**
         * Returns the current recurrence object of an appointment series.
         *
         * This method returns the entire `Recurrence` object for the appointment series.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object. The `value` property of the result is a `Recurrence` object.
         */
        getAsync(callback?: (asyncResult: CommonAPI.AsyncResult<Recurrence>) => void): void;
        /**
         * Sets the recurrence pattern of an appointment series.
         *
         * **Note**: `setAsync` should only be available for series items and not instance items.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `InvalidEndTime`: The appointment end time is before its start time.
         *
         * @param recurrencePattern - A recurrence object.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        setAsync(recurrencePattern: Recurrence, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Sets the recurrence pattern of an appointment series.
         *
         * **Note**: `setAsync` should only be available for series items and not instance items.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `InvalidEndTime`: The appointment end time is before its start time.
         *
         * @param recurrencePattern - A recurrence object.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        setAsync(recurrencePattern: Recurrence, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * Provides updated recurrence object that raised the `Office.EventType.RecurrenceChanged` event.
     *
     * @remarks
     * [Api set: Mailbox 1.7]
     */
    export interface RecurrenceChangedEventArgs {
        /**
         * Gets the updated recurrence object.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         */
        recurrence: Recurrence;
        /**
         * Gets the type of the event. For details, refer to {@link https://learn.microsoft.com/javascript/api/office/office.eventtype | Office.EventType}.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         */
        type: "olkRecurrenceChanged";
    }
    /**
     * Represents the properties of the recurrence.
     *
     * @remarks
     * [Api set: Mailbox 1.7]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface RecurrenceProperties {
        /**
         * Represents the period between instances of the same recurring series.
         */
        interval: number;
        /**
         * Represents the day of the month.
         */
        dayOfMonth?: number;
        /**
         * Represents the day of the week or type of day, for example, weekend day vs weekday.
         */
        dayOfWeek?: MailboxEnums.Days | string;
        /**
         * Represents the set of days for this recurrence. Valid values are: 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', and 'Sun'.
         */
        days?: MailboxEnums.Days[] | string[];
        /**
         * Represents the number of the week in the selected month e.g., 'first' for first week of the month.
         */
        weekNumber?: MailboxEnums.WeekNumber | string;
        /**
         * Represents the month.
         */
        month?: MailboxEnums.Month | string;
        /**
         * Represents your chosen first day of the week otherwise the default is the value in the current user's settings.
         * Valid values are: 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', and 'Sun'.
         */
        firstDayOfWeek?: MailboxEnums.Days | string;
    }
    /**
     * Represents the time zone of the recurrence.
     *
     * @remarks
     * [Api set: Mailbox 1.7]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface RecurrenceTimeZone {
        /**
         * Represents the name of the recurrence time zone.
         */
        name: MailboxEnums.RecurrenceTimeZone | string;

        /**
         * Integer value representing the difference in minutes between the local time zone and UTC at the date that the meeting series began.
         */
        offset?: number;
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
         *
         * **Important**: This link must be publicly accessible, without need for authentication by Exchange Online servers. However, with
         * on-premises Exchange, the link can be accessible on a private network as long as it doesn't need further authentication.
         */
        url?: string;
        /**
         * Only used if type is set to file. If true, indicates that the attachment will be shown inline in the message body, and should not be
         * displayed in the attachment list.
         */
        inLine?: boolean;
        /**
         * Only used if type is set to item. The EWS item ID of the attachment. This is a string up to 100 characters.
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
         * An array of {@link Office.ReplyFormAttachment | ReplyFormAttachment} that are either file or item attachments.
         */
        attachments?: ReplyFormAttachment[];
        /**
         * When the reply display call completes, the function passed in the callback parameter is called with a single parameter,
         * `asyncResult`, which is an `Office.AsyncResult` object.
         */
        callback?: (asyncResult: CommonAPI.AsyncResult<any>) => void;
        /**
         * An object literal that contains the following property:-
         * `asyncContext`: Developers can provide any object they wish to access in the callback function.
         */
        options?: CommonAPI.AsyncContextOptions;
    }
    /**
     * The settings created by using the methods of the `RoamingSettings` object are saved per add-in and per user.
     * That is, they are available only to the add-in that created them, and only from the user's mailbox in which they are saved.
     *
     * While the Outlook add-in API limits access to these settings to only the add-in that created them, these settings shouldn't be considered
     * secure storage. They can be accessed by Exchange Web Services or Extended MAPI.
     * They shouldn't be used to store sensitive information, such as user credentials or security tokens.
     *
     * The name of a setting is a String, while the value can be a String, Number, Boolean, null, Object, or Array.
     *
     * The `RoamingSettings` object is accessible via the `roamingSettings` property in the `Office.context` namespace.
     * 
     * To learn more about `RoamingSettings`, see
     * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/metadata-for-an-outlook-add-in | Get and set add-in metadata for an Outlook add-in}.
     *
     * @remarks
     * [Api set: Mailbox 1.1]
     *
     * **Important**:
     *
     * - The `RoamingSettings` object is initialized from the persisted storage only when the add-in is first loaded.
     * For task panes, this means that it's only initialized when the task pane first opens.
     * If the task pane navigates to another page or reloads the current page, the in-memory object is reset to its initial values, even if
     * your add-in has persisted changes.
     * The persisted changes will not be available until the task pane (or item in the case of UI-less add-ins) is closed and reopened.
     *
     * - When set and saved through Outlook on Windows or Mac, these settings are reflected in Outlook on the web only after a browser refresh.
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface RoamingSettings {
        /**
         * Retrieves the specified setting.
         *
         * @returns Type: String | Number | Boolean | Object | Array
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param name - The case-sensitive name of the setting to retrieve.
         */
        get(name: string): any;
        /**
         * Removes the specified setting.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param name - The case-sensitive name of the setting to remove.
         */
        remove(name: string): void;
        /**
         * Saves the settings.
         *
         * Any settings previously saved by an add-in are loaded when it's initialized, so during the lifetime of the session you can just use
         * the set and get methods to work with the in-memory copy of the settings property bag.
         * When you want to persist the settings so that they're available the next time the add-in is used, use the `saveAsync` method.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`.
         */
        saveAsync(callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Sets or creates the specified setting.
         *
         * The `set` method creates a new setting of the specified name if it doesn't already exist, or sets an existing setting of the specified name.
         * The value is stored in the document as the serialized JSON representation of its data type.
         *
         * A maximum of 32KB is available for the settings of each add-in. An error with code 9057 is thrown when that size limit is exceeded.
         *
         * Any changes made to settings using the `set` method will not be saved to the server until the `saveAsync` method is called.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **restricted**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * @param name - The case-sensitive name of the setting to set or create.
         * @param value - Specifies the value to be stored.
         */
        set(name: string, value: any): void;
    }
    /**
     * Provides methods to get and set the sensitivity level of an appointment. To learn more about sensitivity levels, see
     * {@link https://support.microsoft.com/office/4a76d05b-6c29-4a0d-9096-71784a6b12c1 | Mark your email as Normal, Personal, Private, or Confidential}.
     *
     * @remarks
     * [Api set: Mailbox preview]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
     *
     * @beta
     */
    export interface Sensitivity {
        /**
         * Gets the sensitivity level of an appointment.
         *
         * @remarks
         * [Api set: Mailbox preview]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`,
         *                 which is an `Office.AsyncResult` object. The sensitivity level of the appointment is returned in the `asyncResult.value` property.
         *
         * @beta
         */
        getAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<MailboxEnums.AppointmentSensitivityType>) => void): void;
        /**
         * Gets the sensitivity level of an appointment.
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`,
         *                 which is an `Office.AsyncResult` object. The sensitivity level of the appointment is returned in the `asyncResult.value` property.
         *
         * @remarks
         * [Api set: Mailbox preview]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @beta
         */
        getAsync(callback: (asyncResult: CommonAPI.AsyncResult<MailboxEnums.AppointmentSensitivityType>) => void): void;
        /**
         * Sets the sensitivity level of an appointment.
         *
         * @remarks
         * [Api set: Mailbox preview]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param sensitivity - The sensitivity level as an enum or string.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         *
         * @beta
         */
        setAsync(sensitivity: MailboxEnums.AppointmentSensitivityType | string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Sets the sensitivity level of an appointment.
         *
         * @remarks
         * [Api set: Mailbox preview]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param sensitivity - The sensitivity level as an enum or string.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         *
         * @beta
         */
        setAsync(sensitivity: MailboxEnums.AppointmentSensitivityType | string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * Provides methods to get or set the sensitivity label of a message or appointment. For more information on sensitivity labels, see
     * {@link https://learn.microsoft.com/microsoft-365/compliance/sensitivity-labels | Learn about sensitivity labels}.
     * 
     * @remarks
     * [Api set: Mailbox preview]
     * 
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
     *
     * **Important**: To use the sensitivity label feature in your add-in, you must have a Microsoft 365 E5 subscription.
     *
     * To learn more about how to manage sensitivity labels in your add-in, see
     * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/sensitivity-label | Manage the sensitivity label of your message or appointment in compose mode (preview)}.
     *
     * @beta
     */
    export interface SensitivityLabel {
        /**
         * Gets the unique identifier (GUID) of the sensitivity label applied to a message or appointment being composed.
         * 
         * @remarks
         * [Api set: Mailbox preview]
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Important**: To use the sensitivity label feature in your add-in, you must have a Microsoft 365 E5 subscription.
         *
         * To learn more about how to manage sensitivity labels in your add-in, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/sensitivity-label | Manage the sensitivity label of your message or appointment in compose mode (preview)}.
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`,
         *                 which is an `Office.AsyncResult` object. The sensitivity label's GUID is returned in the
         *                 `asyncResult.value` property.
         *
         * @beta
         */
        getAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Gets the unique identifier (GUID) of the sensitivity label applied to a message or appointment being composed.
         * 
         * @remarks
         * [Api set: Mailbox preview]
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Important**: To use the sensitivity label feature in your add-in, you must have a Microsoft 365 E5 subscription.
         *
         * To learn more about how to manage sensitivity labels in your add-in, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/sensitivity-label | Manage the sensitivity label of your message or appointment in compose mode (preview)}.
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`,
         *                 which is an `Office.AsyncResult` object. The sensitivity label's GUID is returned in the
         *                 `asyncResult.value` property.
         *
         * @beta
         */
        getAsync(callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Applies the specified sensitivity label to the message or appointment being composed.
         * 
         * @remarks
         * [Api set: Mailbox preview]
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Important**: To use the sensitivity label feature in your add-in, you must have a Microsoft 365 E5 subscription.
         *
         * **Tip**: To determine the sensitivity labels available for use, call the `Office.context.sensitivityLabelsCatalog.getAsync` method.
         *
         * To learn more about how to manage sensitivity labels in your add-in, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/sensitivity-label | Manage the sensitivity label of your message or appointment in compose mode (preview)}.
         * 
         * @param sensitivityLabel - The sensitivity label to be applied to the message or appointment being composed. The parameter value can be a sensitivity label's
         *                         unique identifier (GUID) or a {@link Office.SensitivityLabelDetails | SensitivityLabelDetails} object.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`,
         *                 which is an `Office.AsyncResult` object.
         *
         * @beta
         */
        setAsync(sensitivityLabel: string | SensitivityLabelDetails, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Applies the specified sensitivity label to the message or appointment being composed.
         * 
         * @remarks
         * [Api set: Mailbox preview]
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Important**: To use the sensitivity label feature in your add-in, you must have a Microsoft 365 E5 subscription.
         *
         * **Tip**: To determine the sensitivity labels available for use, call the `Office.context.sensitivityLabelsCatalog.getAsync` method.
         *
         * To learn more about how to manage sensitivity labels in your add-in, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/sensitivity-label | Manage the sensitivity label of your message or appointment in compose mode (preview)}.
         * 
         * @param sensitivityLabel - The sensitivity label to be applied to the message or appointment being composed. The parameter value can be a sensitivity label's
         *                         unique identifier (GUID) or a {@link Office.SensitivityLabelDetails | SensitivityLabelDetails} object.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`,
         *                 which is an `Office.AsyncResult` object.
         * 
         * @beta
         */
        setAsync(sensitivityLabel: string | SensitivityLabelDetails, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * Provides the change status of the sensitivity label applied to a message or appointment in compose mode. This information is provided when the
     * `Office.EventType.SensitivityLabelChanged` event is raised.
     * 
     * @remarks
     * [Api set: Mailbox preview]
     *
     * **Important**: To use the sensitivity label feature in your add-in, you must have a Microsoft 365 E5 subscription.
     *
     * To learn more about how to manage sensitivity labels in your add-in, see
     * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/sensitivity-label | Manage the sensitivity label of your message or appointment in compose mode (preview)}.
     *
     * @beta
     */
    export interface SensitivityLabelChangedEventArgs {
        /**
         * The type of event that was raised. For details, refer to {@link https://learn.microsoft.com/javascript/api/office/office.eventtype | Office.EventType}.
         * 
         * @remarks
         * [Api set: Mailbox preview]
         * 
         * @beta
         */
        type: "olkSensitivityLabelChanged";
    }
    /**
     * Represents the properties of available sensitivity labels in Outlook.
     * 
     * @remarks
     * [Api set: Mailbox preview]
     * 
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
     *
     * **Important**: To use the sensitivity label feature in your add-in, you must have a Microsoft 365 E5 subscription.
     *
     * To learn more about how to manage sensitivity labels in your add-in, see
     * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/sensitivity-label | Manage the sensitivity label of your message or appointment in compose mode (preview)}.
     * 
     * @beta
     */
    export interface SensitivityLabelDetails {
        /**
         * The color of the sensitivity label.
         * 
         * @beta
         */
        color: string;
        /**
         * The {@link https://learn.microsoft.com/microsoft-365/compliance/sensitivity-labels#sublabels-grouping-labels | sublabels} of the sensitivity label.
         * Returns `null` if a label doesn't have any sublabels.
         * 
         * @beta
         */
        children: SensitivityLabelDetails[];
        /**
         * The unique identifier (GUID) of the sensitivity label.
         * 
         * @beta
         */
        id: string;
        /**
         * The name of the sensitivity label.
         * 
         * @beta
         */
        name: string;
        /**
         * The description of the sensitivity label.
         * 
         * @beta
         */
        tooltip: string;
    }
    /**
     * Provides methods to check the status of the catalog of {@link https://learn.microsoft.com/microsoft-365/compliance/sensitivity-labels | sensitivity labels}
     * in Outlook and retrieve all available sensitivity labels if the catalog is enabled.
     * 
     * @remarks
     * [Api set: Mailbox preview]
     * 
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
     *
     * **Important**: To use the sensitivity label feature in your add-in, you must have a Microsoft 365 E5 subscription.
     *
     * To learn more about how to manage sensitivity labels in your add-in, see
     * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/sensitivity-label | Manage the sensitivity label of your message or appointment in compose mode (preview)}.
     *
     * @beta
     */
    export interface SensitivityLabelsCatalog {
        /**
         * Gets all the sensitivity labels that are enabled in Outlook.
         * 
         * @remarks
         * [Api set: Mailbox preview]
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Important**: To use the sensitivity label feature in your add-in, you must have a Microsoft 365 E5 subscription.
         *
         * **Recommended**: To determine whether the catalog of sensitivity labels is enabled in Outlook, call `getIsEnabledAsync` before using `getAsync`.
         *
         * To learn more about how to manage sensitivity labels in your add-in, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/sensitivity-label | Manage the sensitivity label of your message or appointment in compose mode (preview)}.
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`,
         *                 which is an `Office.AsyncResult` object. The available sensitivity labels and their properties are returned in the
         *                 `asyncResult.value` property.
         *  
         * @beta
         */
        getAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<SensitivityLabelDetails[]>) => void): void;
        /**
         * Gets all the sensitivity labels that are enabled in Outlook.
         * 
         * @remarks
         * [Api set: Mailbox preview]
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Important**: To use the sensitivity label feature in your add-in, you must have a Microsoft 365 E5 subscription.
         *
         * **Recommended**: To determine whether the catalog of sensitivity labels is enabled in Outlook, call `getIsEnabledAsync` before using `getAsync`.
         *
         * To learn more about how to manage sensitivity labels in your add-in, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/sensitivity-label | Manage the sensitivity label of your message or appointment in compose mode (preview)}.
         * 
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`,
         *                 which is an `Office.AsyncResult` object. The available sensitivity labels and their properties are returned in the
         *                 `asyncResult.value` property.
         *  
         * @beta
         */
        getAsync(callback: (asyncResult: CommonAPI.AsyncResult<SensitivityLabelDetails[]>) => void): void;
        /**
         * Checks whether the catalog of sensitivity labels is enabled in Outlook.
         * 
         * @remarks
         * [Api set: Mailbox preview]
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         * 
         * **Important**: The catalog of sensitivity labels is configured by an organization's administrator. For more information, see
         * {@link https://learn.microsoft.com/microsoft-365/compliance/get-started-with-sensitivity-labels | Get started with sensitivity labels}.
         *
         * **Important**: To use the sensitivity label feature in your add-in, you must have a Microsoft 365 E5 subscription.
         *
         * To learn more about how to manage sensitivity labels in your add-in, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/sensitivity-label | Manage the sensitivity label of your message or appointment in compose mode (preview)}.
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`,
         *                 which is an `Office.AsyncResult` object. The status of the catalog of sensitivity labels is returned in the `asyncResult.value` property.
         * 
         * @beta
         */
        getIsEnabledAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<boolean>) => void): void;
        /**
         * Checks whether the catalog of sensitivity labels is enabled in Outlook.
         * 
         * @remarks
         * [Api set: Mailbox preview]
         * 
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Important**: The catalog of sensitivity labels is configured by an organization's administrator. For more information, see
         * {@link https://learn.microsoft.com/microsoft-365/compliance/get-started-with-sensitivity-labels | Get started with sensitivity labels}.
         *
         * **Important**: To use the sensitivity label feature in your add-in, you must have a Microsoft 365 E5 subscription.
         *
         * To learn more about how to manage sensitivity labels in your add-in, see
         * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/sensitivity-label | Manage the sensitivity label of your message or appointment in compose mode (preview)}.
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`,
         *                 which is an `Office.AsyncResult` object. The status of the catalog of sensitivity labels is returned in the `asyncResult.value` property.
         * 
         * @beta
         */
        getIsEnabledAsync(callback: (asyncResult: CommonAPI.AsyncResult<boolean>) => void): void;
    }
    /**
     * The `SeriesTime` object provides methods to get and set the dates and times of appointments in a recurring series and get the dates and times
     * of meeting requests in a recurring series.
     *
     * @remarks
     * [Api set: Mailbox 1.7]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface SeriesTime {
        /**
         * Gets the duration in minutes of a usual instance in a recurring appointment series.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        getDuration(): number;
        /**
         * Gets the end date of a recurrence pattern in the following
         * {@link https://www.iso.org/iso-8601-date-and-time-format.html | ISO 8601} date format: "YYYY-MM-DD".
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        getEndDate(): string;
        /**
         * Gets the end time of a usual appointment or meeting request instance of a recurrence pattern in whichever time zone that the user or
         * add-in set the recurrence pattern using the following {@link https://www.iso.org/iso-8601-date-and-time-format.html | ISO 8601} format:
         * "THH:mm:ss:mmm".
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        getEndTime(): string;
        /**
         * Gets the start date of a recurrence pattern in the following
         * {@link https://www.iso.org/iso-8601-date-and-time-format.html | ISO 8601} date format: "YYYY-MM-DD".
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        getStartDate(): string;
        /**
         * Gets the start time of a usual appointment instance of a recurrence pattern in whichever time zone that the user/add-in set the
         * recurrence pattern using the following {@link https://www.iso.org/iso-8601-date-and-time-format.html | ISO 8601} format: "THH:mm:ss:mmm".
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        getStartTime(): string;
        /**
         * Sets the duration of all appointments in a recurrence pattern. This will also change the end time of the recurrence pattern.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param minutes - The length of the appointment in minutes.
         */
        setDuration(minutes: number): void;
        /**
         * Sets the end date of a recurring appointment series.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param year - The year value of the end date.
         * @param month - The month value of the end date. Valid range is 0-11 where 0 represents the 1st month and 11 represents the 12th month.
         * @param day - The day value of the end date.
         */
        setEndDate(year: number, month: number, day: number): void;
        /**
         * Sets the end date of a recurring appointment series.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param date - End date of the recurring appointment series represented in the
         * {@link https://www.iso.org/iso-8601-date-and-time-format.html | ISO 8601} date format: "YYYY-MM-DD".
         */
        setEndDate(date: string): void;
        /**
         * Sets the start date of a recurring appointment series.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param year - The year value of the start date.
         * @param month - The month value of the start date. Valid range is 0-11 where 0 represents the 1st month and 11 represents the 12th month.
         * @param day - The day value of the start date.
         */
        setStartDate(year:number, month:number, day:number): void;
        /**
         * Sets the start date of a recurring appointment series.
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param date - Start date of the recurring appointment series represented in the
         * {@link https://www.iso.org/iso-8601-date-and-time-format.html | ISO 8601} date format: "YYYY-MM-DD".
         */
        setStartDate(date:string): void;
        /**
         * Sets the start time of all instances of a recurring appointment series in whichever time zone the recurrence pattern is set
         * (the item's time zone is used by default).
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param hours - The hour value of the start time. Valid range: 0-24.
         * @param minutes - The minute value of the start time. Valid range: 0-59.
         */
        setStartTime(hours: number, minutes: number): void;
        /**
         * Sets the start time of all instances of a recurring appointment series in whichever time zone the recurrence pattern is set
         * (the item's time zone is used by default).
         *
         * @remarks
         * [Api set: Mailbox 1.7]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param time - Start time of all instances represented by standard datetime string format: "THH:mm:ss:mmm".
         */
        setStartTime(time: string): void;
    }
    /**
     * Provides methods to manage an item's session data.
     *
     * **Important**: The entire SessionData object is limited to 50,000 characters per add-in.
     *
     * @remarks
     * [Api set: Mailbox 1.11]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
     */
    export interface SessionData {
        /**
         * Clears all session data key-value pairs.
         *
         * @remarks
         * [Api set: Mailbox 1.11]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        clearAsync(options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Clears all session data key-value pairs.
         *
         * @remarks
         * [Api set: Mailbox 1.11]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        clearAsync(callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Gets all session data key-value pairs.
         *
         * @remarks
         * [Api set: Mailbox 1.11]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        getAllAsync(callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Gets the session data value of the specified key.
         *
         * @remarks
         * [Api set: Mailbox 1.11]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param name - The session data key.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        getAsync(name: string, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Removes a session data key-value pair.
         *
         * @remarks
         * [Api set: Mailbox 1.11]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param name - The session data key.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        removeAsync(name: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Removes a session data key-value pair.
         *
         * @remarks
         * [Api set: Mailbox 1.11]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param name - The session data key.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter,
         *                `asyncResult`, which is an `Office.AsyncResult` object.
         */
        removeAsync(name: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Sets a session data key-value pair.
         *
         * **Important**: The entire SessionData object is limited to 50,000 characters per add-in.
         *
         * @remarks
         * [Api set: Mailbox 1.11]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose

         * @param name - The session data key.
         * @param value - The session data value as a string.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        setAsync(name: string, value: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Sets a session data key-value pair.
         *
         * **Important**: The entire SessionData object is limited to 50,000 characters per add-in.
         *
         * @remarks
         * [Api set: Mailbox 1.11]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose

         * @param name - The session data key.
         * @param value - The session data value as a string.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *                 type `Office.AsyncResult`.
         */
        setAsync(name: string, value: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * Represents the properties of an appointment or message in a shared folder or shared mailbox (now in preview).
     *
     * For more information on how this object is used, see the
     * {@link https://learn.microsoft.com/office/dev/add-ins/outlook/delegate-access | shared folders and shared mailbox} article.
     *
     * @remarks
     * [Api set: Mailbox 1.8]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface SharedProperties {
        /**
         * The email address of the owner of a shared item.
         */
        owner: string;
        /**
         * The REST API's base URL (currently https://outlook.office.com/api).
         *
         * Use with `targetMailbox` to construct the REST operation's URL.
         *
         * Example usage: `targetRestUrl + "/{api_version}/users/" + targetMailbox + "/{REST_operation}"`
         */
        targetRestUrl: string;
        /**
         * The location of the owner's mailbox for the delegate's access. This location may differ based on the Outlook client.
         *
         * Use with `targetRestUrl` to construct the REST operation's URL.
         *
         * Example usage: `targetRestUrl + "/{api_version}/users/" + targetMailbox + "/{REST_operation}"`
         */
        targetMailbox: string;
        /**
         * The permissions that the delegate has on a shared folder, or the user has on a shared mailbox.
         */
        delegatePermissions: MailboxEnums.DelegatePermissions;
    }
    /**
     * Provides methods to get and set the subject of an appointment or message in an Outlook add-in.
     *
     * @remarks
     * [Api set: Mailbox 1.1]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
     */
    export interface Subject {
        /**
         * Gets the subject of an appointment or message.
         *
         * The `getAsync` method starts an asynchronous call to the Exchange server to get the subject of an appointment or message.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. The `value` property of the result is the subject of the item.
         */
        getAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Gets the subject of an appointment or message.
         *
         * The `getAsync` method starts an asynchronous call to the Exchange server to get the subject of an appointment or message.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. The `value` property of the result is the subject of the item.
         */
        getAsync(callback: (asyncResult: CommonAPI.AsyncResult<string>) => void): void;
        /**
         * Sets the subject of an appointment or message.
         *
         * The `setAsync` method starts an asynchronous call to the Exchange server to set the subject of an appointment or message.
         * Setting the subject overwrites the current subject, but leaves any prefixes, such as "Fwd:" or "Re:" in place.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `DataExceedsMaximumSize`: The subject parameter is longer than 255 characters.
         *
         * @param subject - The subject of the appointment or message. The string is limited to 255 characters.
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. If setting the subject fails, the `asyncResult.error` property will contain an error code.
         */
        setAsync(subject: string, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Sets the subject of an appointment or message.
         *
         * The `setAsync` method starts an asynchronous call to the Exchange server to set the subject of an appointment or message.
         * Setting the subject overwrites the current subject, but leaves any prefixes, such as "Fwd:" or "Re:" in place.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `DataExceedsMaximumSize`: The subject parameter is longer than 255 characters.
         *
         * @param subject - The subject of the appointment or message. The string is limited to 255 characters.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                 of type `Office.AsyncResult`. If setting the subject fails, the `asyncResult.error` property will contain an error code.
         */
        setAsync(subject: string, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * Represents a suggested task identified in an item. Read mode only.
     *
     * The list of tasks suggested in an email message is returned in the `taskSuggestions` property of the {@link Office.Entities | Entities} object
     * that is returned when the `getEntities` or `getEntitiesByType` method is called on the active item.
     *
     * @remarks
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Read
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
     * The `Time` object is returned as the start or end property of an appointment in compose mode.
     *
     * @remarks
     * [Api set: Mailbox 1.1]
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
     */
    export interface Time {
        /**
         * Gets the start or end time of an appointment.
         *
         * The date and time is provided as a `Date` object in the `asyncResult.value` property. The value is in Coordinated Universal Time (UTC).
         * You can convert the UTC time to the local client time by using the `convertToLocalClientTime` method.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - When the method completes, the function passed in the `callback` parameter is called with a single parameter
         *                  of type `Office.AsyncResult`. The `value` property of the result is a `Date` object.
         */
        getAsync(options: CommonAPI.AsyncContextOptions, callback: (asyncResult: CommonAPI.AsyncResult<Date>) => void): void;
        /**
         * Gets the start or end time of an appointment.
         *
         * The date and time is provided as a `Date` object in the `asyncResult.value` property. The value is in Coordinated Universal Time (UTC).
         * You can convert the UTC time to the local client time by using the `convertToLocalClientTime` method.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * @param callback - When the method completes, the function passed in the callback parameter is called with a single parameter
         *                  of type `Office.AsyncResult`. The `value` property of the result is a `Date` object.
         */
        getAsync(callback: (asyncResult: CommonAPI.AsyncResult<Date>) => void): void;
        /**
         * Sets the start or end time of an appointment.
         *
         * If the `setAsync` method is called on the start property, the `end` property will be adjusted to maintain the duration of the appointment as
         * previously set. If the `setAsync` method is called on the `end` property, the duration of the appointment will be extended to the new end time.
         *
         * The time must be in UTC; you can get the correct UTC time by using the `convertToUtcClientTime` method.
         *
         * **Important**: In the Windows client, you can't use this method to update the start or end of a recurrence.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `InvalidEndTime`: The appointment end time is before the appointment start time.
         *
         * @param dateTime - A date-time object in Coordinated Universal Time (UTC).
         * @param options - An object literal that contains one or more of the following properties:-
         *        `asyncContext`: Developers can provide any object they wish to access in the callback function.
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *               type `Office.AsyncResult`. If setting the date and time fails, the `asyncResult.error` property will contain an error code.
         */
        setAsync(dateTime: Date, options: CommonAPI.AsyncContextOptions, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
        /**
         * Sets the start or end time of an appointment.
         *
         * If the `setAsync` method is called on the start property, the `end` property will be adjusted to maintain the duration of the appointment as
         * previously set. If the `setAsync` method is called on the `end` property, the duration of the appointment will be extended to the new end time.
         *
         * The time must be in UTC; you can get the correct UTC time by using the `convertToUtcClientTime` method.
         *
         * **Important**: In the Windows client, you can't use this method to update the start or end of a recurrence.
         *
         * @remarks
         * [Api set: Mailbox 1.1]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read/write item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose
         *
         * **Errors**:
         *
         * - `InvalidEndTime`: The appointment end time is before the appointment start time.
         *
         * @param dateTime - A date-time object in Coordinated Universal Time (UTC).
         * @param callback - Optional. When the method completes, the function passed in the `callback` parameter is called with a single parameter of
         *               type `Office.AsyncResult`. If setting the date and time fails, the `asyncResult.error` property will contain an error code.
         */
        setAsync(dateTime: Date, callback?: (asyncResult: CommonAPI.AsyncResult<void>) => void): void;
    }
    /**
     * Information about the user associated with the mailbox. This includes their account type, display name, email address, and time zone.
     *
     * @remarks
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
     *
     * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
     */
    export interface UserProfile {
        /**
         * Gets the account type of the user associated with the mailbox.
         *
         * **Note**: This member is currently only supported in Outlook 2016 or later on Mac, build 16.9.1212 and greater.
         *
         * @remarks
         * [Api set: Mailbox 1.6]
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         *
         * The possible account types are listed in the following table.
         *
         * <table>
         *   <tr>
         *     <th>Value</th>
         *     <th>Description</th>
         *   </tr>
         *   <tr>
         *     <td>enterprise</td>
         *     <td>The mailbox is on an on-premises Exchange server.</td>
         *   </tr>
         *   <tr>
         *     <td>gmail</td>
         *     <td>The mailbox is associated with a Gmail account.</td>
         *   </tr>
         *   <tr>
         *     <td>office365</td>
         *     <td>The mailbox is associated with a Microsoft 365 work or school account.</td>
         *   </tr>
         *   <tr>
         *     <td>outlookCom</td>
         *     <td>The mailbox is associated with a personal Outlook.com account.</td>
         *   </tr>
         * </table>
         * 
         * **Note**: For hybrid Exchange environments, the returned account type value depends on where the mailbox is hosted.
         * If the mailbox is on an on-premises server, the account type value is **enterprise**. However, if it's hosted on 
         * Exchange Online, the account type value is **office365**.
         */
        accountType: string;
        /**
         * Gets the user's display name.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        displayName: string;
        /**
         * Gets the user's SMTP email address.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        emailAddress: string;
        /**
         * Gets the user's time zone in Windows format.
         *
         * The system time zone is usually returned. However, in Outlook on the web, the default time zone in the calendar preferences is returned instead.
         *
         * @remarks
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions | Minimum permission level}**: **read item**
         *
         * **{@link https://learn.microsoft.com/office/dev/add-ins/outlook/outlook-add-ins-overview#extension-points | Applicable Outlook mode}**: Compose or Read
         */
        timeZone: string;
    }
}


////////////////////////////////////////////////////////////////
/////////////////////// End Exchange APIs //////////////////////
////////////////////////////////////////////////////////////////