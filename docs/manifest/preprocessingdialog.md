---
title: PreProcessingDialog element in the manifest file
description: The PreProcessingDialog element configures the preprocessing dialog of a spam-reporting add-in in Outlook.
ms.date: 05/20/2024
ms.localizationpriority: medium
---

# PreProcessingDialog element

Configures the preprocessing dialog of a spam-reporting add-in in Outlook, so that users can provide additional information about the message they're reporting.

To learn more about how to implement the spam reporting feature in your add-in, see [Implement an integrated spam-reporting add-in](/office/dev/add-ins/outlook/spam-reporting).

**Add-in type**: Mail

**Valid only in these VersionOverrides schemas**:

- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [Mailbox 1.14](../requirement-sets/outlook/requirement-set-1.14/outlook-requirement-set-1.14.md)

## Contained in

- [ReportPhishingCustomization](reportphishingcustomization.md)

## Attributes

None.

## Child elements

| Element | Required | Description |
| :------ | :------: | :------ |
| **Title** | Yes | Specifies the custom title of the preprocessing dialog. Its **resid** attribute must be set to the value of the **id** attribute of a [String](string.md) in the [ShortStrings](shortstrings.md) element under the [Resources](resources.md) element. |
| **Description** | Yes | Specifies the custom text that appears in the preprocessing dialog. Its **resid** attribute must be set to the value of the **id** attribute of a [String](string.md) in the [LongStrings](longstrings.md) element under the [Resources](resources.md) element. |
| [ReportingOptions](reportingoptions.md) | No | Lists up to five options a user can select from the preprocessing dialog to provide a reason for reporting a message. |
| **FreeTextLabel** | No | Adds a text box to the preprocessing dialog to allow users to provide additional information on the message they're reporting. Its **resid** attribute sets the title of the text box. The **resid** attribute must be set to the value of the **id** attribute of a [String](string.md) in the [ShortStrings](shortstrings.md) element under the [Resources](resources.md) element. |
| [MoreInfo](moreinfo.md) | No | Specifies the custom text and URL to provide informational resources to the users. The custom text and URL configured in this element appear below the text provided in the **\<Description\>** element. |

## Example

```xml
<PreProcessingDialog>
  <Title resid="PreProcessingDialog.Label"/>
  <Description resid="PreProcessingDialog.Description"/>
  <ReportingOptions>
    <Title resid="OptionsTitle.Label"/>
    <Option resid="Option1.Label"/>
    <Option resid="Option2.Label"/>
    <Option resid="Option3.Label"/>
    <Option resid="Option4.Label"/>
  </ReportingOptions>
  <FreeTextLabel resid="FreeText.Label"/>
  <MoreInfo>
    <MoreInfoText resid="MoreInfo.Label"/>
    <MoreInfoUrl resid="MoreInfo.Url"/>
  </MoreInfo>
</PreProcessingDialog>
```
