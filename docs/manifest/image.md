---
title: Image element in the manifest file
description: The Image element enables you to specify the URL of an image used for an icon.
ms.date: 02/07/2023
ms.localizationpriority: medium
---

# Image element

Provides the URL of an image file that is used as an icon.

**Add-in type:** Task pane, Mail

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.0
- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

> [!NOTE]
> You must use Secure Sockets Layer (SSL) for all URLs in the **\<Image\>** element and any child **\<Override\>** elements.

Each icon must have three **\<Image\>** elements, one for each of the three mandatory sizes:

- 16x16
- 32x32
- 80x80

The following additional sizes are also supported, but not required.

- 20x20
- 24x24
- 40x40
- 48x48
- 64x64

The following image file formats are supported.

- BMP
- EXIF
- GIF
- JPG
- PNG
- TIFF

> [!IMPORTANT]
>
> - If the image is your add-in's representative icon, see [Create effective listings in AppSource and within Office](/partner-center/marketplace-offers/create-effective-office-store-listings#create-an-icon-for-your-add-in) for size and other requirements.
> - Office Add-ins require the ability to cache image resources for performance purposes. For this reason, the server hosting an image resource must not add any CACHE-CONTROL directives to the response header. These directives result in Office automatically substituting a generic or default image. To force the use of new icons on your development computer, [Clear the Office cache](/office/dev/add-ins/testing/clear-cache). To force the use of new icons on your end-user's computers, you must give the new icons different URLs from the old ones.

## Child elements

|  Element |  Type  |  Description  |
|:-----|:-----:|:-----|
|  [Override](override.md)           |  image   |  Provides a way to override the URL depending on a specified locale. |

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
