---
title: EquivalentAddins element in the manifest file
description: Specifies backwards compatibility with an equivalent COM add-in, XLL, or both.
ms.date: 01/04/2022
ms.localizationpriority: medium
---

# EquivalentAddins element

Specifies backwards compatibility with an equivalent COM add-in, XLL, or both.

[!INCLUDE [Support note for equivalent add-ins feature](/office/dev/add-ins/includes/equivalent-add-in-support-note)]

**Add-in type:** Task pane, Mail, Custom function

**Valid only in these VersionOverrides schemas**:

- Task pane 1.0
- Mail 1.1

For more information, see [Version overrides in the manifest](/office/dev/add-ins/develop/add-in-manifests#version-overrides-in-the-manifest).

## Syntax

```XML
<EquivalentAddins>
...  
</EquivalentAddins>  
```

## Contained in

[VersionOverrides](versionoverrides.md)

## Must contain

[EquivalentAddin](equivalentaddin.md)

## See also

- [Make your custom functions compatible with XLL user-defined functions](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf)
- [Make your Office Add-in compatible with an existing COM add-in](/office/dev/add-ins/develop/make-office-add-in-compatible-with-existing-com-add-in)