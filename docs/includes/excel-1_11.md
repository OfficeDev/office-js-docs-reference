| Class | Fields | Description |
|:---|:---|:---|
|[Application](/.application)|[cultureInfo](/.application#excel-javascript/api/excel/-application-cultureinfo-member)|Provides information based on current system culture settings.|
||[decimalSeparator](/.application#excel-javascript/api/excel/-application-decimalseparator-member)|Gets the string used as the decimal separator for numeric values.|
||[thousandsSeparator](/.application#excel-javascript/api/excel/-application-thousandsseparator-member)|Gets the string used to separate groups of digits to the left of the decimal for numeric values.|
||[useSystemSeparators](/.application#excel-javascript/api/excel/-application-usesystemseparators-member)|Specifies if the system separators of Excel are enabled.|
|[Comment](/.comment)|[mentions](/.comment#excel-javascript/api/excel/-comment-mentions-member)|Gets the entities (e.g., people) that are mentioned in comments.|
||[resolved](/.comment#excel-javascript/api/excel/-comment-resolved-member)|The comment thread status.|
||[richContent](/.comment#excel-javascript/api/excel/-comment-richcontent-member)|Gets the rich comment content (e.g., mentions in comments).|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/.comment#excel-javascript/api/excel/-comment-updatementions-member(1))|Updates the comment content with a specially formatted string and a list of mentions.|
|[CommentCollection](/.commentcollection)|[add(cellAddress: Range \| string, content: CommentRichContent \| string, contentType?: Excel.ContentType)](/.commentcollection#excel-javascript/api/excel/-commentcollection-add-member(1))|Creates a new comment with the given content on the given cell.|
|[CommentMention](/.commentmention)|[email](/.commentmention#excel-javascript/api/excel/-commentmention-email-member)|The email address of the entity that is mentioned in a comment.|
||[id](/.commentmention#excel-javascript/api/excel/-commentmention-id-member)|The ID of the entity.|
||[name](/.commentmention#excel-javascript/api/excel/-commentmention-name-member)|The name of the entity that is mentioned in a comment.|
|[CommentReply](/.commentreply)|[mentions](/.commentreply#excel-javascript/api/excel/-commentreply-mentions-member)|The entities (e.g., people) that are mentioned in comments.|
||[resolved](/.commentreply#excel-javascript/api/excel/-commentreply-resolved-member)|The comment reply status.|
||[richContent](/.commentreply#excel-javascript/api/excel/-commentreply-richcontent-member)|The rich comment content (e.g., mentions in comments).|
||[updateMentions(contentWithMentions: Excel.CommentRichContent)](/.commentreply#excel-javascript/api/excel/-commentreply-updatementions-member(1))|Updates the comment content with a specially formatted string and a list of mentions.|
|[CommentReplyCollection](/.commentreplycollection)|[add(content: CommentRichContent \| string, contentType?: Excel.ContentType)](/.commentreplycollection#excel-javascript/api/excel/-commentreplycollection-add-member(1))|Creates a comment reply for a comment.|
|[CommentRichContent](/.commentrichcontent)|[mentions](/.commentrichcontent#excel-javascript/api/excel/-commentrichcontent-mentions-member)|An array containing all the entities (e.g., people) mentioned within the comment.|
||[richContent](/.commentrichcontent#excel-javascript/api/excel/-commentrichcontent-richcontent-member)|Specifies the rich content of the comment (e.g., comment content with mentions, the first mentioned entity has an ID attribute of 0, and the second mentioned entity has an ID attribute of 1).|
|[CultureInfo](/.cultureinfo)|[name](/.cultureinfo#excel-javascript/api/excel/-cultureinfo-name-member)|Gets the culture name in the format languagecode2-country/regioncode2 (e.g., "zh-cn" or "en-us").|
||[numberFormat](/.cultureinfo#excel-javascript/api/excel/-cultureinfo-numberformat-member)|Defines the culturally appropriate format of displaying numbers.|
|[NumberFormatInfo](/.numberformatinfo)|[numberDecimalSeparator](/.numberformatinfo#excel-javascript/api/excel/-numberformatinfo-numberdecimalseparator-member)|Gets the string used as the decimal separator for numeric values.|
||[numberGroupSeparator](/.numberformatinfo#excel-javascript/api/excel/-numberformatinfo-numbergroupseparator-member)|Gets the string used to separate groups of digits to the left of the decimal for numeric values.|
|[Range](/.range)|[moveTo(destinationRange: Range \| string)](/.range#excel-javascript/api/excel/-range-moveto-member(1))|Moves cell values, formatting, and formulas from current range to the destination range, replacing the old information in those cells.|
|[RangeFormat](/.rangeformat)|[adjustIndent(amount: number)](/.rangeformat#excel-javascript/api/excel/-rangeformat-adjustindent-member(1))|Adjusts the indentation of the range formatting.|
|[Workbook](/.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/.workbook#excel-javascript/api/excel/-workbook-close-member(1))|Close current workbook.|
||[save(saveBehavior?: Excel.SaveBehavior)](/.workbook#excel-javascript/api/excel/-workbook-save-member(1))|Save current workbook.|
|[Worksheet](/.worksheet)|[onRowHiddenChanged](/.worksheet#excel-javascript/api/excel/-worksheet-onrowhiddenchanged-member)|Occurs when the hidden state of one or more rows has changed on a specific worksheet.|
|[WorksheetCalculatedEventArgs](/.worksheetcalculatedeventargs)|[address](/.worksheetcalculatedeventargs#excel-javascript/api/excel/-worksheetcalculatedeventargs-address-member)|The address of the range that completed calculation.|
|[WorksheetCollection](/.worksheetcollection)|[onRowHiddenChanged](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-onrowhiddenchanged-member)|Occurs when the hidden state of one or more rows has changed on a specific worksheet.|
|[WorksheetRowHiddenChangedEventArgs](/.worksheetrowhiddenchangedeventargs)|[address](/.worksheetrowhiddenchangedeventargs#excel-javascript/api/excel/-worksheetrowhiddenchangedeventargs-address-member)|Gets the range address that represents the changed area of a specific worksheet.|
||[changeType](/.worksheetrowhiddenchangedeventargs#excel-javascript/api/excel/-worksheetrowhiddenchangedeventargs-changetype-member)|Gets the type of change that represents how the event was triggered.|
||[source](/.worksheetrowhiddenchangedeventargs#excel-javascript/api/excel/-worksheetrowhiddenchangedeventargs-source-member)|Gets the source of the event.|
||[type](/.worksheetrowhiddenchangedeventargs#excel-javascript/api/excel/-worksheetrowhiddenchangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.worksheetrowhiddenchangedeventargs#excel-javascript/api/excel/-worksheetrowhiddenchangedeventargs-worksheetid-member)|Gets the ID of the worksheet in which the data changed.|
