---
title: ReportingOptions element in the manifest file (preview)
description: The ReportingOptions element specifies the reporting options listed in the preprocessing dialog of a spam-reporting add-in in Outlook.
ms.date: 07/14/2023
ms.localizationpriority: medium
---

# ReportingOptions element (preview)

Specifies the reporting options listed in the preprocessing dialog of a spam-reporting add-in in Outlook.

To learn more about how to implement the spam reporting feature in your add-in, see [Implement an integrated spam-reporting add-in (preview)](/office/dev/add-ins/outlook/spam-reporting).

**Add-in type**: Mail

**Valid only in these VersionOverrides schemas**:

- Mail 1.1

For more information, see [Version overrides in the manifest](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [Mailbox preview](../requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview.md)

## Contained in

- [PreProcessingDialog (preview)](preprocessingdialog.md)

## Attributes

None.

## Child elements

| Element | Required | Description |
| :------ | :------: | :------ |
| **Title** | Yes | Specifies the custom title of the preprocessing dialog. Its **resid** attribute must be set to the value of the **id** attribute of a [String](string.md) in the [ShortStrings](shortstrings.md) element under the [Resources](resources.md) element. |
| **Option** | Yes | Specifies a custom option that a user can select from the preprocessing dialog to provide a reason for reporting a message. You can add up to *five* options, but must set at least one option. |

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
