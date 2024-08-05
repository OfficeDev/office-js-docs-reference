---
title: String element in the manifest file
description: The String element enables you to specify the text of a Label, Title, or Description element.
ms.date: 07/20/2023
ms.localizationpriority: medium
---

# String element

Defines a string that can be used as the text of one or more **\<Description\>**, **\<Label\>**, or **\<Title\>** elements.

> [!NOTE]
>
> - When the parent element is **\<ShortStrings\>**, the value can be no more than 125 characters.
> - When the parent element is **\<LongStrings\>**, the value can be no more than 250 characters.

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

## Attributes

| Attribute | Required | Description |
| :----- | :-----: | :----- |
| **id** | Yes | Specifies the unique identifier of a string resource. |
| **DefaultValue** | Yes | Specifies the custom text of a resource. If the parent element is **\<ShortStrings\>**, the text can have a maximum of 125 characters. However, if its parent element is **\<LongStrings\>**, the text can have a maximum of 250 characters. |

## Child elements

| Element | Type | Description |
|:-----|:-----:|:-----|
| [Override](override.md) | string | Provides a way to override the string depending on a specified locale. |

## Example

```XML
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_32x32" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp32-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_80x80" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp80-icon_default.png" />
        </bt:Image>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residDesktopFuncUrl" DefaultValue="https://www.contoso.com/Pages/Home.aspx">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Pages/Home.aspx" />
        </bt:Url>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residLabel" DefaultValue="GetData">
          <bt:Override Locale="ja-jp" Value="JA-JP-GetData" />
        </bt:String>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residToolTip" DefaultValue="Get data for your document.">
          <bt:Override Locale="ja-jp" Value="JA-JP - Get data for your document." />
        </bt:String>
      </bt:LongStrings>
    </Resources>
```
