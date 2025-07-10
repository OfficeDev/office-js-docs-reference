---
title: Type element in the manifest file
description: The Type element specifies if the equivalent add-in is a COM or VSTO add-in or an XLL.
ms.date: 07/12/2025
ms.localizationpriority: medium
---

# Type element

Specifies if the equivalent add-in is a COM or VSTO add-in or an XLL.

**Add-in type:** Task pane, Custom function

## Syntax

```XML
    <Type> [COM | XLL] </Type>  
```

## Contained in

- [EquivalentAddin](equivalentaddin.md)

## Add-in type values

You must specify one of the following values for the `Type` element.

- COM: Specifies the equivalent add-in is a COM or a VSTO add-in.
- XLL: Specifies the equivalent add-in is an Excel XLL.

> [!IMPORTANT]
> Use `COM` as the value of the `Type` element for both COM and VSTO add-ins.

## See also

- [Make your custom functions compatible with XLL user-defined functions](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf)
- [Make your Office Add-in compatible with an existing COM or VSTO add-in](/office/dev/add-ins/develop/make-office-add-in-compatible-with-existing-com-add-in)