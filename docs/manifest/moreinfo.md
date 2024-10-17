---
title: MoreInfo element in the manifest file
description: The MoreInfo element specifies the custom text and URL that direct users to informational resources from the preprocessing dialog of a spam-reporting add-in in Outlook.
ms.date: 06/10/2024
ms.localizationpriority: medium
---

# MoreInfo element

Specifies custom text and URL to provide informational resources to the users from the preprocessing dialog of a spam-reporting add-in in Outlook. The information provided in this element helps users identify and report unsolicited messages. It appears after the text provided in the [Description](preprocessingdialog.md#child-elements) child element of the **\<PreProcessingDialog\>** element.

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

None.

## Child elements

| Element | Required | Description |
| ------- | ------- | -------|
| **MoreInfoText** | Yes | Specifies additional information in the preprocessing dialog of a spam-reporting add-in to help users report unsolicited messages. The **resid** attribute must be set to the value of the **id** attribute of a [String](string.md) in the [ShortStrings](shortstrings.md) element under the [Resources](resources.md) element. Depending on the Outlook client, the text specified in the **\<MoreInfoText\>** element appears before the URL that's provided in the **\<MoreInfoUrl\>** element or as link text for the URL. For more information, see [MoreInfoText](#moreinfotext).|
| **MoreInfoUrl** | Yes | Specifies the URL of a site containing informational resources in the preprocessing dialog of a spam-reporting add-in. Its **resid** attribute must be set to the value of the **id** attribute of a [Url](url.md) in the [Urls](urls.md) element under the [Resources](resources.md) element. |

### MoreInfoText

In Outlook on the web, classic Outlook on Windows (starting in Version 2404 (Build 17526.15020)), and [new Outlook on Windows](https://support.microsoft.com/office/656bb8d9-5a60-49b2-a98b-ba7822bc7627), the **\<MoreInfoText\>** element specifies the custom link text of the URL that's provided in the **\<MoreInfoUrl\>** element. On these Outlook clients, the link is automatically prefixed with the static string, "For more info go to: ".

:::image type="content" source="images/outlook-spam-processing-dialog.png" alt-text="Sample preprocessing dialog of a spam-reporting add-in in Outlook on the web and supported versions of Outlook on Windows (classic and new). The link specified in the **\<MoreInfo\>** element is prepended with the static text, 'For more info go to \:'.":::

In Outlook on Mac and earlier supported versions of classic Outlook on Windows (prior to Version 2404 (Build 17526.15020)), the **\<MoreInfoText\>** element specifies the custom text that appears before the bare URL that's provided in the **\<MoreInfoUrl\>** element.

:::image type="content" source="images/outlook-spam-processing-dialog-mac.png" alt-text="Sample preprocessing dialog of a spam-reporting add-in in Outlook on Mac and earlier supported versions of classic Outlook on Windows. The custom text specified in the **\<MoreInfoText\>** element appears before the bare URL specified in the **\<MoreInfoUrl\>** element.":::

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
