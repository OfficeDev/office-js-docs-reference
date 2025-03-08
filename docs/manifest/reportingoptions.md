---
title: ReportingOptions element in the manifest file
description: The ReportingOptions element specifies the reporting options listed in the preprocessing dialog of a spam-reporting add-in in Outlook.
ms.date: 03/11/2025
ms.localizationpriority: medium
---

# ReportingOptions element

Specifies the reporting options listed in the preprocessing dialog of a spam-reporting add-in in Outlook.

To learn more about how to implement the spam reporting feature in your add-in, see [Implement an integrated spam-reporting add-in](/office/dev/add-ins/outlook/spam-reporting).

**Add-in type**: Mail

**Valid only in these VersionOverrides schemas**:

- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [Mailbox 1.14](../requirement-sets/outlook/requirement-set-1.14/outlook-requirement-set-1.14.md)

## Contained in

- [PreProcessingDialog](preprocessingdialog.md)

## Attributes

| Attribute | Required | Description |
|:-----|:-----:|:-----|
| **inputType** | No | Specifies the input type of the reporting options in the preprocessing dialog. If the `inputType` attribute isn't included, the reporting options appear as checkboxes. To use radio buttons, set the `inputType` attribute to `Radio`. You can only use one input type in the dialog.<br><br>**Important**: The **inputType** attribute was introduced in [requirement set 1.15](../requirement-sets/outlook/requirement-set-1.15/outlook-requirement-set-1.15.md). Learn more about its [supported clients and platforms](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets). |

## Child elements

| Element | Required | Description |
| :------ | :------: | :------ |
| **Title** | Yes | Specifies a custom title for the reporting options in the preprocessing dialog. Its **resid** attribute must be set to the value of the **id** attribute of a [String](string.md) in the [ShortStrings](shortstrings.md) element under the [Resources](resources.md) element. |
| **Option** | Yes | Specifies a custom option that a user can select from the preprocessing dialog to provide a reason for reporting a message. You can add up to *five* options, but must specify at least one option. |

## Example

```xml
<PreProcessingDialog>
  ...
  <ReportingOptions>
    <Title resid="OptionsTitle.Label"/>
    <Option resid="Option1.Label"/>
    <Option resid="Option2.Label"/>
    <Option resid="Option3.Label"/>
    <Option resid="Option4.Label"/>
  </ReportingOptions>
  ...
</PreProcessingDialog>
```
