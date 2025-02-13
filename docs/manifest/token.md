---
title: Token element in the manifest file
description: Specifies a token or wildcard that can be used with URL templates in the manifest.
ms.date: 02/12/2025
ms.localizationpriority: medium
---


# Token element

Defines an individual URL token.

**Add-in type:** Task pane

## Syntax

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## Contained in

- [Tokens](tokens.md)

## Can contain

The **\<Token\>** element can contain the following child element depending on the add-in type.

|Element|Content|Mail|TaskPane|
|:-----|:-----:|:-----:|:-----:|
|[Override](override.md)|No|No|Yes|

## Attributes

|Attribute|Description|
|:-----|:-----|
|DefaultValue|Default value for this token if no condition in any child **\<Override\>** element matches.|
|Name|Token name. This name is user-defined. The type of the token is determined by the type attribute.|
|xsi:type|Defines the kind of Token. This attribute should be set to one of:  `"RequirementsToken"`,  or  `"LocaleToken"`.|

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