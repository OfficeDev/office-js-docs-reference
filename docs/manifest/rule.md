---
title: Rule element in the manifest file
description: The Rule element specifies the activation rules that should be evaluated for this contextual mail add-in.
ms.date: 09/04/2024
ms.localizationpriority: medium
---

# Rule element

Specifies the activation rules that should be evaluated for this contextual mail add-in.

**Add-in type:** Mail (contextual)

[!INCLUDE [outlook-contextual-add-ins-retirement](../includes/outlook-contextual-add-ins-retirement.md)]

## Contained in

- [OfficeApp](officeapp.md)
- [ExtensionPoint](extensionpoint.md) ([**CustomPane** (obsolete)](https://devblogs.microsoft.com/microsoft365dev/make-your-add-ins-available-in-the-office-ribbon/), [DetectedEntity](extensionpoint.md#detectedentity))

## Attributes

| Attribute | Required | Description |
|:-----|:-----:|:-----|
| **xsi:type** | Yes | The type of rule being defined. |

The type of rule can be one of the following:

- [ItemIs](#itemis-rule)
- [ItemHasAttachment](#itemhasattachment-rule)
- [ItemHasKnownEntity](#itemhasknownentity-rule-deprecated) (deprecated)
- [ItemHasRegularExpressionMatch](#itemhasregularexpressionmatch-rule)
- [RuleCollection](#rulecollection)

## ItemIs rule

Defines a rule that evaluates to true if the selected item in read or compose form is of the specified type. You can specify more than one `ItemIs` rule in the manifest.

### Attributes

| Attribute | Required | Description |
|:-----|:-----:|:-----|
| **ItemType** | Yes | Specifies the item type to match. Can be `Message` or `Appointment`. `Message` item type includes email, meeting requests, meeting responses, and meeting cancellations. |
| **FormType** | No (within [ExtensionPoint](extensionpoint.md)), Yes (within [OfficeApp](officeapp.md)) | Specifies whether the app should appear in read or edit form for the item. Can be one of the following: `Read`, `Edit`, `ReadOrEdit`. If specified on a `Rule` within an `ExtensionPoint`, this value MUST be `Read`. |
| **ItemClass** | No | Specifies the custom message class to match. For more information, see [Item Types and Message Classes](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes). |
| **IncludeSubClasses** | No | Specifies whether the rule should evaluate to true if the item is of a subclass of the specified message class; the default is `false`. |

#### ItemType

The following table describes the `Appointment` and `Message` item types you can specify in the `ItemType` attribute.

| Value | Description |
|:-----|:-----|
| **Appointment** | Specifies an item in an Outlook calendar. This includes a meeting item that has been responded to and has an organizer and attendees, or an appointment that doesn't have an organizer or attendee and is simply an item on the calendar. This corresponds to the IPM.Appointment message class in Outlook. |
| **Message** | Specifies one of the following items received in typically the Inbox. <ul><li><p>An email message. This corresponds to the IPM.Note message class in Outlook.</p></li><li><p>A meeting request, response, or cancellation. This corresponds to the following message classes in Outlook.</p><p>IPM.Schedule.Meeting.Request</p><p>IPM.Schedule.Meeting.Neg</p><p>IPM.Schedule.Meeting.Pos</p><p>IPM.Schedule.Meeting.Tent</p><p>IPM.Schedule.Meeting.Canceled</p></li></ul> |

### Example

```XML
<Rule xsi:type="ItemIs" ItemType= "Message" />
```

## ItemHasAttachment rule

Defines a rule that evaluates to true if the item in read form contains an attachment.

### Example

```XML
<Rule xsi:type="ItemHasAttachment" />
```

## ItemHasKnownEntity rule (deprecated)

> [!WARNING]
> The **\<ItemHasKnownEntity\>** element is now deprecated. Specify regular expression rules using the [ItemHasRegularExpressionMatch](#itemhasregularexpressionmatch-rule) element instead.

Defines a rule that evaluates to true if the item contains text of the specified entity type in its subject or body.

### Attributes

| Attribute | Required | Description |
|:-----|:-----:|:-----|
| **EntityType** | Yes | Specifies the type of entity that must be found for the rule to evaluate to true. Can be one of the following: `MeetingSuggestion`, `TaskSuggestion`, `Address`, `Url`, `PhoneNumber`, `EmailAddress`, or `Contact`. |
| **RegExFilter** | No | Specifies a regular expression to run against this entity for activation. |
| **FilterName** | No | Specifies the name of the regular expression filter, so that it is subsequently possible to refer to it in your add-in's code. |
| **IgnoreCase** | No | Specifies whether to ignore case when matching the regular expression specified by the **RegExFilter** attribute. |
| **Highlight** | No | **Note:** this only applies to **\<Rule\>** elements within **\<ExtensionPoint\>** elements. Specifies how the client should highlight matching entities. Can be one of the following: `all` or `none`. If not specified, the default value is `all`. |

### Example

```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="EmailAddress" />
```

## ItemHasRegularExpressionMatch rule

Defines a rule that evaluates to true if a match for the specified regular expression can be found in the specified property of the item in read form.

To learn how to implement the `ItemHasRegularExpressionMatch` rule, see [Contextual Outlook add-ins](/office/dev/add-ins/outlook/contextual-outlook-add-ins).

### Attributes

| Attribute | Required | Description |
|:-----|:-----:|:-----|
| **RegExName** | Yes | Specifies the name of the regular expression, so that you can refer to the expression in the code for your add-in. |
| **RegExValue** | Yes | Specifies the regular expression that will be evaluated to determine whether the mail add-in should be shown. |
| **PropertyName** | Yes | Specifies the name of the property that the regular expression will be evaluated against. Can be one of the following: `Subject`, `BodyAsPlaintext`, `BodyAsHTML`, or `SenderSMTPAddress`.<br/><br/>If you specify `BodyAsHTML`, Outlook only applies the regular expression if the item body is HTML. Otherwise, Outlook returns no matches for that regular expression.<br/><br/>If you specify `BodyAsPlaintext`, Outlook always applies the regular expression on the item body.<br/><br/>**Important:** If you need to specify the **Highlight** attribute for the **\<Rule\>** element, you must set the **PropertyName** attribute to `BodyAsPlaintext`. |
| **IgnoreCase** | No | Specifies whether to ignore case when matching the regular expression specified by the **RegExName** attribute. |
| **Highlight** | No | Specifies how the client should highlight matching text. This attribute can only be applied to **\<Rule\>** elements within **\<ExtensionPoint\>** elements. Can be one of the following: `all` or `none`. If not specified, the default value is `all`.<br/><br/>**Important:** To specify the **Highlight** attribute in the **\<Rule\>** element, you must set the **PropertyName** attribute to `BodyAsPlaintext`. |

### Example

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="SupportArticleNumber" RegExValue="(\W|^)kb\d{6}(\W|$)" PropertyName="BodyAsPlaintext" IgnoreCase="true" Highlight="all" />
```

## RuleCollection

Defines a collection of rules and the logical operator to use when evaluating them.

### Attributes

| Attribute | Required | Description |
|:-----|:-----:|:-----|
| **Mode** | Yes | Specifies the logical operator to use when evaluating this rule collection. Can be either: `And` or `Or`. |

### Example

```XML
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="Context.Label"/>
  <SourceLocation resid="DetectedEntity.URL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="videoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="BodyAsPlaintext"/>
  </Rule>
</ExtensionPoint>
```

## See also

- [Contextual Outlook add-ins](/office/dev/add-ins/outlook/contextual-outlook-add-ins)
