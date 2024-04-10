---
title: PreProcessingDialog element in the manifest file (preview)
description: The PreProcessingDialog element configures the preprocessing dialog of a spam-reporting add-in in Outlook.
ms.date: 04/11/2024
ms.localizationpriority: medium
---

# PreProcessingDialog element (preview)

Configures the preprocessing dialog of a spam-reporting add-in in Outlook, so that users can provide additional information about the message they're reporting.

To learn more about how to implement the spam reporting feature in your add-in, see [Implement an integrated spam-reporting add-in (preview)](/office/dev/add-ins/outlook/spam-reporting).

**Add-in type**: Mail

**Valid only in these VersionOverrides schemas**:

- Mail 1.1

For more information, see [Version overrides in the manifest](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [Mailbox preview](../requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview.md)

## Contained in

- [ReportPhishingCustomization (preview)](reportphishingcustomization.md)

## Attributes

None.

## Child elements

| Element | Required | Description |
| :------ | :------: | :------ |
| **Title** | Yes | Specifies the custom title of the preprocessing dialog. Its **resid** attribute must be set to the value of the **id** attribute of a [String](string.md) in the [ShortStrings](shortstrings.md) element under the [Resources](resources.md) element. |
| **Description** | Yes | Specifies the custom text that appears in the preprocessing dialog. Its **resid** attribute must be set to the value of the **id** attribute of a [String](string.md) in the [LongStrings](longstrings.md) element under the [Resources](resources.md) element. |
| [ReportingOptions](reportingoptions.md) | No | Lists up to five options a user can select from the preprocessing dialog to provide a reason for reporting a message. |
| **FreeTextLabel** | No | Adds a text box to the preprocessing dialog to allow users to provide additional information on the message they're reporting. Its **resid** attribute sets the title of the text box. The **resid** attribute must be set to the value of the **id** attribute of a [String](string.md) in the [ShortStrings](shortstrings.md) element under the [Resources](resources.md) element. |
| [MoreInfo](moreinfo.md) | No | Specifies a link with custom text and URL to provide informational resources to the users. In the preprocessing dialog, the link is automatically prefixed with "For more info go to: ". It appears below the text provided in the **\<Description\>** element. |

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
