---
title: MoreInfo element in the manifest file (preview)
description: The MoreInfo element specifies the custom text and URL that direct users to informational resources from the pre-processing dialog of a spam reporting add-in in Outlook.
ms.date: 06/22/2023
ms.localizationpriority: medium
---

# MoreInfo element (preview)

Specifies the custom text and URL that direct users to informational resources from the pre-processing dialog of a spam reporting add-in in Outlook. The information provided in this element assists users with identifying and reporting unsolicited messages.

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
| ------- | ------- | -------|
| **MoreInfoText** | Yes | Specifies additional information in the pre-processing dialog of a spam reporting add-in to assist users with reporting unsolicited messages. Its **resid** attribute must be set to the value of the **id** attribute of a [String](string.md) in the [ShortStrings](shortstrings.md) element under the [Resources](resources.md) element. |
| **MoreInfoUrl** | Yes | Specifies the URL of a site containing informational resources in the pre-processing dialog of a spam reporting add-in. Its **resid** attribute must be set to the value of the **id** attribute of a [Url](url.md) in the [Urls](urls.md) element under the [Resources](resources.md) element. |

## Example

```xml
<PreProcessingDialog>
  ...
  <MoreInfo>
    <MoreInfoText resid="MoreInfo.Label"/>
    <MoreInfoUrl resid="MoreInfo.Url"/>
  </MoreInfo>
</PreProcessingDialog>
```
