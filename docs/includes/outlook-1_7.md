| Class | Fields | Description |
|:---|:---|:---|
|[AppointmentCompose](/javascript/api/outlook/office.appointmentcompose)|[addHandlerAsync(eventType: Office.EventType \| string, handler: any, callback?: (asyncResult: Office.AsyncResult\<void\>) => void)](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-addhandlerasync-member(1))|Adds an event handler for a supported event.|
||[addHandlerAsync(eventType: Office.EventType \| string, handler: any, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult\<void\>) => void)](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-addhandlerasync-member(1))|Adds an event handler for a supported event.|
||[organizer](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-organizer-member)|Gets the organizer for the specified meeting.|
||[recurrence](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-recurrence-member)|Gets or sets the recurrence pattern of an appointment.|
||[removeHandlerAsync(eventType: Office.EventType \| string, callback?: (asyncResult: Office.AsyncResult\<void\>) => void)](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-removehandlerasync-member(1))|Removes the event handlers for a supported event type.|
||[removeHandlerAsync(eventType: Office.EventType \| string, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult\<void\>) => void)](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-removehandlerasync-member(1))|Removes the event handlers for a supported event type.|
||[seriesId](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-seriesid-member)|Gets the ID of the series that an instance belongs to.|
|[AppointmentRead](/javascript/api/outlook/office.appointmentread)|[addHandlerAsync(eventType: Office.EventType \| string, handler: any, callback?: (asyncResult: Office.AsyncResult\<void\>) => void)](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-addhandlerasync-member(1))|Adds an event handler for a supported event.|
||[addHandlerAsync(eventType: Office.EventType \| string, handler: any, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult\<void\>) => void)](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-addhandlerasync-member(1))|Adds an event handler for a supported event.|
||[recurrence](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-recurrence-member)|Gets the recurrence pattern of an appointment.|
||[removeHandlerAsync(eventType: Office.EventType \| string, callback?: (asyncResult: Office.AsyncResult\<void\>) => void)](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-removehandlerasync-member(1))|Removes the event handlers for a supported event type.|
||[removeHandlerAsync(eventType: Office.EventType \| string, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult\<void\>) => void)](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-removehandlerasync-member(1))|Removes the event handlers for a supported event type.|
||[seriesId](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-seriesid-member)|Gets the ID of the series that an instance belongs to.|
|[AppointmentTimeChangedEventArgs](/javascript/api/outlook/office.appointmenttimechangedeventargs)|[end](/javascript/api/outlook/office.appointmenttimechangedeventargs#outlook-office-appointmenttimechangedeventargs-end-member)|Gets the appointment end date and time.|
||[start](/javascript/api/outlook/office.appointmenttimechangedeventargs#outlook-office-appointmenttimechangedeventargs-start-member)|Gets the appointment start date and time.|
||[type](/javascript/api/outlook/office.appointmenttimechangedeventargs#outlook-office-appointmenttimechangedeventargs-type-member)|Gets the type of the event.|
|[Days](/javascript/api/outlook/office.mailboxenums.days)|Day|Day of week.|
||Fri|Friday|
||Mon|Monday|
||Sat|Saturday|
||Sun|Sunday|
||Thu|Thursday|
||Tue|Tuesday|
||Wed|Wednesday|
||Weekday|Week day (excludes weekend days): 'Mon', 'Tue', 'Wed', 'Thu', and 'Fri'.|
||Weekendday|Weekend day: 'Sat' and 'Sun'.|
|[From](/javascript/api/outlook/office.from)|[getAsync(callback?: (asyncResult: Office.AsyncResult\<EmailAddressDetails\>) => void)](/javascript/api/outlook/office.from#outlook-office-from-getasync-member(1))|Gets the from value of a message.|
||[getAsync(options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult\<EmailAddressDetails\>) => void)](/javascript/api/outlook/office.from#outlook-office-from-getasync-member(1))|Gets the from value of a message.|
|[MessageCompose](/javascript/api/outlook/office.messagecompose)|[addHandlerAsync(eventType: Office.EventType \| string, handler: any, callback?: (asyncResult: Office.AsyncResult\<void\>) => void)](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-addhandlerasync-member(1))|Adds an event handler for a supported event.|
||[addHandlerAsync(eventType: Office.EventType \| string, handler: any, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult\<void\>) => void)](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-addhandlerasync-member(1))|Adds an event handler for a supported event.|
||[from](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-from-member)|Gets the email address of the sender of a message.|
||[removeHandlerAsync(eventType: Office.EventType \| string, callback?: (asyncResult: Office.AsyncResult\<void\>) => void)](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-removehandlerasync-member(1))|Removes the event handlers for a supported event type.|
||[removeHandlerAsync(eventType: Office.EventType \| string, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult\<void\>) => void)](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-removehandlerasync-member(1))|Removes the event handlers for a supported event type.|
||[seriesId](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-seriesid-member)|Gets the ID of the series that an instance belongs to.|
|[MessageRead](/javascript/api/outlook/office.messageread)|[addHandlerAsync(eventType: Office.EventType \| string, handler: any, callback?: (asyncResult: Office.AsyncResult\<void\>) => void)](/javascript/api/outlook/office.messageread#outlook-office-messageread-addhandlerasync-member(1))|Adds an event handler for a supported event.|
||[addHandlerAsync(eventType: Office.EventType \| string, handler: any, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult\<void\>) => void)](/javascript/api/outlook/office.messageread#outlook-office-messageread-addhandlerasync-member(1))|Adds an event handler for a supported event.|
||[recurrence](/javascript/api/outlook/office.messageread#outlook-office-messageread-recurrence-member)|Gets the recurrence pattern of an appointment.|
||[removeHandlerAsync(eventType: Office.EventType \| string, callback?: (asyncResult: Office.AsyncResult\<void\>) => void)](/javascript/api/outlook/office.messageread#outlook-office-messageread-removehandlerasync-member(1))|Removes the event handlers for a supported event type.|
||[removeHandlerAsync(eventType: Office.EventType \| string, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult\<void\>) => void)](/javascript/api/outlook/office.messageread#outlook-office-messageread-removehandlerasync-member(1))|Removes the event handlers for a supported event type.|
||[seriesId](/javascript/api/outlook/office.messageread#outlook-office-messageread-seriesid-member)|Gets the ID of the series that an instance belongs to.|
|[Month](/javascript/api/outlook/office.mailboxenums.month)|Apr|April|
||Aug|August|
||Dec|December|
||Feb|February|
||Jan|January|
||Jul|July|
||Jun|June|
||Mar|March|
||May|May|
||Nov|November|
||Oct|October|
||Sep|September|
|[Organizer](/javascript/api/outlook/office.organizer)|[getAsync(callback?: (asyncResult: Office.AsyncResult\<EmailAddressDetails\>) => void)](/javascript/api/outlook/office.organizer#outlook-office-organizer-getasync-member(1))|Gets the organizer value of an appointment as an EmailAddressDetails object|
||[getAsync(options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult\<EmailAddressDetails\>) => void)](/javascript/api/outlook/office.organizer#outlook-office-organizer-getasync-member(1))|Gets the organizer value of an appointment as an EmailAddressDetails object|
|[RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs)|[changedRecipientFields](/javascript/api/outlook/office.recipientschangedeventargs#outlook-office-recipientschangedeventargs-changedrecipientfields-member)|Gets an object that indicates change state of recipients fields.|
||[type](/javascript/api/outlook/office.recipientschangedeventargs#outlook-office-recipientschangedeventargs-type-member)|Gets the type of the event.|
|[RecipientsChangedFields](/javascript/api/outlook/office.recipientschangedfields)|[bcc](/javascript/api/outlook/office.recipientschangedfields#outlook-office-recipientschangedfields-bcc-member)|Gets if recipients in the **bcc** field were changed.|
||[cc](/javascript/api/outlook/office.recipientschangedfields#outlook-office-recipientschangedfields-cc-member)|Gets if recipients in the **cc** field were changed.|
||[optionalAttendees](/javascript/api/outlook/office.recipientschangedfields#outlook-office-recipientschangedfields-optionalattendees-member)|Gets if optional attendees were changed.|
||[requiredAttendees](/javascript/api/outlook/office.recipientschangedfields#outlook-office-recipientschangedfields-requiredattendees-member)|Gets if required attendees were changed.|
||[resources](/javascript/api/outlook/office.recipientschangedfields#outlook-office-recipientschangedfields-resources-member)|Gets if resources were changed.|
||[to](/javascript/api/outlook/office.recipientschangedfields#outlook-office-recipientschangedfields-to-member)|Gets if recipients in the **to** field were changed.|
|[Recurrence](/javascript/api/outlook/office.recurrence)|[getAsync(callback?: (asyncResult: Office.AsyncResult\<Recurrence\>) => void)](/javascript/api/outlook/office.recurrence#outlook-office-recurrence-getasync-member(1))|Returns the current recurrence object of an appointment series.|
||[getAsync(options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult\<Recurrence\>) => void)](/javascript/api/outlook/office.recurrence#outlook-office-recurrence-getasync-member(1))|Returns the current recurrence object of an appointment series.|
||[recurrenceProperties](/javascript/api/outlook/office.recurrence#outlook-office-recurrence-recurrenceproperties-member)|Gets or sets the properties of the recurring appointment series.|
||[recurrenceTimeZone](/javascript/api/outlook/office.recurrence#outlook-office-recurrence-recurrencetimezone-member)|Gets or sets the properties of the recurring appointment series.|
||[recurrenceType](/javascript/api/outlook/office.recurrence#outlook-office-recurrence-recurrencetype-member)|Gets or sets the type of the recurring appointment series.|
||[seriesTime](/javascript/api/outlook/office.recurrence#outlook-office-recurrence-seriestime-member)|The SeriesTime object enables you to manage the start and end dates of the recurring appointment series and|
||[setAsync(recurrencePattern: Recurrence, callback?: (asyncResult: Office.AsyncResult\<void\>) => void)](/javascript/api/outlook/office.recurrence#outlook-office-recurrence-setasync-member(1))|Sets the recurrence pattern of an appointment series.|
||[setAsync(recurrencePattern: Recurrence, options: Office.AsyncContextOptions, callback?: (asyncResult: Office.AsyncResult\<void\>) => void)](/javascript/api/outlook/office.recurrence#outlook-office-recurrence-setasync-member(1))|Sets the recurrence pattern of an appointment series.|
|[RecurrenceChangedEventArgs](/javascript/api/outlook/office.recurrencechangedeventargs)|[recurrence](/javascript/api/outlook/office.recurrencechangedeventargs#outlook-office-recurrencechangedeventargs-recurrence-member)|Gets the updated recurrence object.|
||[type](/javascript/api/outlook/office.recurrencechangedeventargs#outlook-office-recurrencechangedeventargs-type-member)|Gets the type of the event.|
|[RecurrenceProperties](/javascript/api/outlook/office.recurrenceproperties)|[dayOfMonth](/javascript/api/outlook/office.recurrenceproperties#outlook-office-recurrenceproperties-dayofmonth-member)|Represents the day of the month.|
||[dayOfWeek](/javascript/api/outlook/office.recurrenceproperties#outlook-office-recurrenceproperties-dayofweek-member)|Represents the day of the week or type of day, for example, weekend day vs weekday.|
||[days](/javascript/api/outlook/office.recurrenceproperties#outlook-office-recurrenceproperties-days-member)|Represents the set of days for this recurrence.|
||[firstDayOfWeek](/javascript/api/outlook/office.recurrenceproperties#outlook-office-recurrenceproperties-firstdayofweek-member)|Represents your chosen first day of the week otherwise the default is the value in the current user's settings.|
||[interval](/javascript/api/outlook/office.recurrenceproperties#outlook-office-recurrenceproperties-interval-member)|Represents the period between instances of the same recurring series.|
||[month](/javascript/api/outlook/office.recurrenceproperties#outlook-office-recurrenceproperties-month-member)|Represents the month.|
||[weekNumber](/javascript/api/outlook/office.recurrenceproperties#outlook-office-recurrenceproperties-weeknumber-member)|Represents the number of the week in the selected month e.g., 'first' for first week of the month.|
|[RecurrenceTimeZone](/javascript/api/outlook/office.recurrencetimezone)|Auscentralstandardtime|Australia Central Standard Time|
||Auseasternstandardtime|AUS Eastern Standard Time|
||Afghanistanstandardtime|Afghanistan Standard Time|
||Alaskanstandardtime|Alaskan Standard Time|
||Aleutianstandardtime|Aleutian Standard Time|
||Altaistandardtime|Altai Standard Time|
||Arabstandardtime|Arab Standard Time|
||Arabianstandardtime|Arabian Standard Time|
||Arabicstandardtime|Arabic Standard Time|
||Argentinastandardtime|Argentina Standard Time|
||Astrakhanstandardtime|Astrakhan Standard Time|
||Atlanticstandardtime|Atlantic Standard Time|
||[](/javascript/api/outlook/office.recurrencetimezone#outlook-office-recurrencetimezone-auscentralw_standardtime-member)|Australia Central West Standard Time|
||Azerbaijanstandardtime|Azerbaijan Standard Time|
||Azoresstandardtime|Azores Standard Time|
||Bahiastandardtime|Bahia Standard Time|
||Bangladeshstandardtime|Bangladesh Standard Time|
||Belarusstandardtime|Belarus Standard Time|
||Bougainvillestandardtime|Bougainville Standard Time|
||Canadacentralstandardtime|Canada Central Standard Time|
||Capeverdestandardtime|Cape Verde Standard Time|
||Caucasusstandardtime|Caucasus Standard Time|
||Cenaustraliastandardtime|Central Australia Standard Time|
||Centralamericastandardtime|Central America Standard Time|
||Centralasiastandardtime|Central Asia Standard Time|
||Centralbrazilianstandardtime|Central Brazilian Standard Time|
||Centraleuropestandardtime|Central Europe Standard Time|
||Centraleuropeanstandardtime|Central European Standard Time|
||Centralpacificstandardtime|Central Pacific Standard Time|
||Centralstandardtime|Central Standard Time|
||[](/javascript/api/outlook/office.recurrencetimezone#outlook-office-recurrencetimezone-centralstandardtime_mexico-member)|Central Standard Time (Mexico)|
||Chathamislandsstandardtime|Chatham Islands Standard Time|
||Chinastandardtime|China Standard Time|
||Cubastandardtime|Cuba Standard Time|
||Datelinestandardtime|Dateline Standard Time|
||[](/javascript/api/outlook/office.recurrencetimezone#outlook-office-recurrencetimezone-e_africastandardtime-member)|East Africa Standard Time|
||[](/javascript/api/outlook/office.recurrencetimezone#outlook-office-recurrencetimezone-e_australiastandardtime-member)|East Australia Standard Time|
||[](/javascript/api/outlook/office.recurrencetimezone#outlook-office-recurrencetimezone-e_europestandardtime-member)|East Europe Standard Time|
||[](/javascript/api/outlook/office.recurrencetimezone#outlook-office-recurrencetimezone-e_southamericastandardtime-member)|East South America Standard Time|
||Easterislandstandardtime|Easter Island Standard Time|
||Easternstandardtime|Eastern Standard Time|
||[](/javascript/api/outlook/office.recurrencetimezone#outlook-office-recurrencetimezone-easternstandardtime_mexico-member)|Eastern Standard Time (Mexico)|
||Egyptstandardtime|Egypt Standard Time|
||Ekaterinburgstandardtime|Ekaterinburg Standard Time|
||Flestandardtime|FLE Standard Time|
||Fijistandardtime|Fiji Standard Time|
||Gmtstandardtime|GMT Standard Time|
||Gtbstandardtime|GTB Standard Time|
||Georgianstandardtime|Georgian Standard Time|
||Greenlandstandardtime|Greenland Standard Time|
||Greenwichstandardtime|Greenwich Standard Time|
||Haitistandardtime|Haiti Standard Time|
||Hawaiianstandardtime|Hawaiian Standard Time|
||Indiastandardtime|India Standard Time|
||Iranstandardtime|Iran Standard Time|
||Israelstandardtime|Israel Standard Time|
||Jordanstandardtime|Jordan Standard Time|
||Kaliningradstandardtime|Kaliningrad Standard Time|
||Kamchatkastandardtime|Kamchatka Standard Time|
||Koreastandardtime|Korea Standard Time|
||Libyastandardtime|Libya Standard Time|
||Lineislandsstandardtime|Line Islands Standard Time|
||Lordhowestandardtime|Lord Howe Standard Time|
||Magadanstandardtime|Magadan Standard Time|
||Magallanesstandardtime|Magallanes Standard Time|
||Marquesasstandardtime|Marquesas Standard Time|
||Mauritiusstandardtime|Mauritius Standard Time|
||Midatlanticstandardtime|Mid-Atlantic Standard Time|
||Middleeaststandardtime|Middle East Standard Time|
||Montevideostandardtime|Montevideo Standard Time|
||Moroccostandardtime|Morocco Standard Time|
||Mountainstandardtime|Mountain Standard Time|
||[](/javascript/api/outlook/office.recurrencetimezone#outlook-office-recurrencetimezone-mountainstandardtime_mexico-member)|Mountain Standard Time (Mexico)|
||Myanmarstandardtime|Myanmar Standard Time|
||[](/javascript/api/outlook/office.recurrencetimezone#outlook-office-recurrencetimezone-n_centralasiastandardtime-member)|North Central Asia Standard Time|
||Namibiastandardtime|Namibia Standard Time|
||Nepalstandardtime|Nepal Standard Time|
||Newzealandstandardtime|New Zealand Standard Time|
||Newfoundlandstandardtime|Newfoundland Standard Time|
||Norfolkstandardtime|Norfolk Standard Time|
||Northasiaeaststandardtime|North Asia East Standard Time|
||Northasiastandardtime|North Asia Standard Time|
||Northkoreastandardtime|North Korea Standard Time|
||Omskstandardtime|Omsk Standard Time|
||Pacificsastandardtime|Pacific SA Standard Time|
||Pacificstandardtime|Pacific Standard Time|
||Pacificstandardtimemexico|Pacific Standard Time (Mexico)|
||Pakistanstandardtime|Pakistan Standard Time|
||Paraguaystandardtime|Paraguay Standard Time|
||Romancestandardtime|Romance Standard Time|
||Russiatimezone10|Russia Time Zone 10|
||Russiatimezone11|Russia Time Zone 11|
||Russiatimezone3|Russia Time Zone 3|
||Russianstandardtime|Russian Standard Time|
||Saeasternstandardtime|SA Eastern Standard Time|
||Sapacificstandardtime|SA Pacific Standard Time|
||Sawesternstandardtime|SA Western Standard Time|
||Seasiastandardtime|Southeast Asia Standard Time|
||Saintpierrestandardtime|Saint Pierre Standard Time|
||Sakhalinstandardtime|Sakhalin Standard Time|
||Samoastandardtime|Samoa Standard Time|
||Saratovstandardtime|Saratov Standard Time|
||Singaporestandardtime|Singapore Standard Time|
||Southafricastandardtime|South Africa Standard Time|
||Srilankastandardtime|Sri Lanka Standard Time|
||Sudanstandardtime|Sudan Standard Time|
||Syriastandardtime|Syria Standard Time|
||Taipeistandardtime|Taipei Standard Time|
||Tasmaniastandardtime|Tasmania Standard Time|
||Tocantinsstandardtime|Tocantins Standard Time|
||Tokyostandardtime|Tokyo Standard Time|
||Tomskstandardtime|Tomsk Standard Time|
||Tongastandardtime|Tonga Standard Time|
||Transbaikalstandardtime|Transbaikal Standard Time|
||Turkeystandardtime|Turkey Standard Time|
||Turksandcaicosstandardtime|Turks And Caicos Standard Time|
||Useasternstandardtime|United States Eastern Standard Time|
||Usmountainstandardtime|United States Mountain Standard Time|
||Utc|Coordinated Universal Time (UTC)|
||Utcminus02|Coordinated Universal Time (UTC) - 2 hours|
||Utcminus08|Coordinated Universal Time (UTC) - 8 hours|
||Utcminus09|Coordinated Universal Time (UTC) - 9 hours|
||Utcminus11|Coordinated Universal Time (UTC) - 11 hours|
||Utcplus12|Coordinated Universal Time (UTC) + 12 hours|
||Utcplus13|Coordinated Universal Time (UTC) + 13 hours|
||Ulaanbaatarstandardtime|Ulaanbaatar Standard Time|
||Venezuelastandardtime|Venezuela Standard Time|
||Vladivostokstandardtime|Vladivostok Standard Time|
||[](/javascript/api/outlook/office.recurrencetimezone#outlook-office-recurrencetimezone-w_australiastandardtime-member)|West Australia Standard Time|
||[](/javascript/api/outlook/office.recurrencetimezone#outlook-office-recurrencetimezone-w_centralafricastandardtime-member)|West Central Africa Standard Time|
||[](/javascript/api/outlook/office.recurrencetimezone#outlook-office-recurrencetimezone-w_europestandardtime-member)|West Europe Standard Time|
||[](/javascript/api/outlook/office.recurrencetimezone#outlook-office-recurrencetimezone-w_mongoliastandardtime-member)|West Mongolia Standard Time|
||Westasiastandardtime|West Asia Standard Time|
||Westbankstandardtime|West Bank Standard Time|
||Westpacificstandardtime|West Pacific Standard Time|
||Yakutskstandardtime|Yakutsk Standard Time|
|[RecurrenceTimeZone](/javascript/api/outlook/office.recurrencetimezone)|[name](/javascript/api/outlook/office.recurrencetimezone#outlook-office-recurrencetimezone-name-member)|Represents the name of the recurrence time zone.|
||[offset](/javascript/api/outlook/office.recurrencetimezone#outlook-office-recurrencetimezone-offset-member)|Integer value representing the difference in minutes between the local time zone and UTC at the date that the meeting series began.|
|[RecurrenceType](/javascript/api/outlook/office.mailboxenums.recurrencetype)|Daily|Daily.|
||Monthly|Monthly.|
||Weekday|Weekday.|
||Weekly|Weekly.|
||Yearly|Yearly.|
|[SeriesTime](/javascript/api/outlook/office.seriestime)|[getDuration()](/javascript/api/outlook/office.seriestime#outlook-office-seriestime-getduration-member(1))|Gets the duration in minutes of a usual instance in a recurring appointment series.|
||[getEndDate()](/javascript/api/outlook/office.seriestime#outlook-office-seriestime-getenddate-member(1))|Gets the end date of a recurrence pattern in the following|
||[getEndTime()](/javascript/api/outlook/office.seriestime#outlook-office-seriestime-getendtime-member(1))|Gets the end time of a usual appointment or meeting request instance of a recurrence pattern in whichever time zone that the user or|
||[getStartDate()](/javascript/api/outlook/office.seriestime#outlook-office-seriestime-getstartdate-member(1))|Gets the start date of a recurrence pattern in the following|
||[getStartTime()](/javascript/api/outlook/office.seriestime#outlook-office-seriestime-getstarttime-member(1))|Gets the start time of a usual appointment instance of a recurrence pattern in whichever time zone that the user/add-in set the|
||[setDuration(minutes: number)](/javascript/api/outlook/office.seriestime#outlook-office-seriestime-setduration-member(1))|Sets the duration of all appointments in a recurrence pattern.|
||[setEndDate(date: string)](/javascript/api/outlook/office.seriestime#outlook-office-seriestime-setenddate-member(1))|Sets the end date of a recurring appointment series.|
||[setEndDate(year: number, month: number, day: number)](/javascript/api/outlook/office.seriestime#outlook-office-seriestime-setenddate-member(1))|Sets the end date of a recurring appointment series.|
||[setStartDate(date:string)](/javascript/api/outlook/office.seriestime#outlook-office-seriestime-setstartdate-member(1))|Sets the start date of a recurring appointment series.|
||[setStartDate(year:number, month:number, day:number)](/javascript/api/outlook/office.seriestime#outlook-office-seriestime-setstartdate-member(1))|Sets the start date of a recurring appointment series.|
||[setStartTime(hours: number, minutes: number)](/javascript/api/outlook/office.seriestime#outlook-office-seriestime-setstarttime-member(1))|Sets the start time of all instances of a recurring appointment series in whichever time zone the recurrence pattern is set|
||[setStartTime(time: string)](/javascript/api/outlook/office.seriestime#outlook-office-seriestime-setstarttime-member(1))|Sets the start time of all instances of a recurring appointment series in whichever time zone the recurrence pattern is set|
|[WeekNumber](/javascript/api/outlook/office.mailboxenums.weeknumber)|First|First week of the month.|
||Fourth|Fourth week of the month.|
||Last|Last week of the month.|
||Second|Second week of the month.|
||Third|Third week of the month.|
