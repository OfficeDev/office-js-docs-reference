---
title: SourceLocation element in the manifest file
description: The SourceLocation element specifies the source file locations for your Office Add-in.
ms.date: 12/01/2023
ms.localizationpriority: medium
---

# SourceLocation element

Specifies the source file locations for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path. 

The source location can be any website, and you can organize the files at that domain in any directory structure you want. Typically, all the files in the add-in, except the manifest, are deployed to the source location. For a single-page application, there is usually one file. You can host image or other supplementary files on a CDN, or other domain, provided that you register the domain with an [AppDomain](appdomain.md) element.

**Add-in type:** Content, Task pane, Mail

## Syntax

```XML
<SourceLocation DefaultValue="string" />
```

## Contained in

- [DefaultSettings](defaultsettings.md) (Content and task pane add-ins)
- [Form](form.md) (Mail add-ins)

## Can contain

- [Override](override.md)

## Attributes

|Attribute|Type|Required|Description|
|:-----|:-----:|:-----:|:-----|
|DefaultValue|URL|Yes|Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.|
