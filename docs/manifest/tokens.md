---
title: Tokens element in the manifest file
description: Specifies tokens or wildcards that can be used with URL templates in the the manifest.
ms.date: 02/12/2025
ms.localizationpriority: medium
---

# Tokens element

Defines tokens that could be used in template URLs.

**Add-in type:** Task pane

## Syntax

```XML
<Tokens></Tokens>
```

## Contained in

- [ExtendedOverrides](extendedoverrides.md)

## Must contain

The **\<Tokens\>** element can contain the following child elements depending on the add-in type.

|Element|Content|Mail|TaskPane|
|:-----|:-----:|:-----:|:-----:|
|[Token](token.md)|No|No|Yes|

## Example

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.locale}/extended-manifest-overrides.json">
    <Tokens>
      <Token Name="locale" DefaultValue="en-us" xsi:type="LocaleToken">
        <Override Locale="es-*" Value="es-es" />
        <Override Locale="es-mx" Value="es-mx" />
        <Override Locale="fr-*" Value="fr-fr" />
        <Override Locale="ja-jp" Value="ja-jp" />
      </Token>
    <Tokens>
  </ExtendedOverrides>
</OfficeApp>
```