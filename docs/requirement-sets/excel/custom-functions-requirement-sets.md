---
title: Custom Functions requirement sets
description: Details about the Custom Functions requirement sets for Excel JavaScript API.
ms.date: 09/14/2022
ms.prod: excel
ms.localizationpriority: medium
---

# Custom Functions requirement sets

[Custom Functions](/office/dev/add-ins/excel/custom-functions-overview) use separate requirement sets from the core Excel JavaScript APIs. The following table lists the Custom Functions requirement sets, the supported Office client applications, and the **minimum** builds or versions for those applications.

| Requirement set | Office on Windows<br>(subscription) | Office on Windows<br>(retail perpetual Office 2016 or later) | Office on Windows<br>(volume-licensed perpetual) | Office on Mac | Office on iPad | Office on the web |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| CustomFunctionsRuntime 1.3 | Build 16.0.13127.20296 | Build 16.0.13127.20296 | Office 2021: Build 16.0.14326.20454 | 16.40.20081000 | Not supported | Supported |
| CustomFunctionsRuntime 1.2 | Build 16.0.14326.20454 | Build 16.0.14326.20454 | Office 2021: Build 16.0.14326.20454 | 16.34.20020900 | Not supported | Supported |
| CustomFunctionsRuntime 1.1 | Build 16.0.12527.20092 | Build 16.0.12527.20092 | Office 2021: Build 16.0.14326.20454 | 16.34 | Not supported | Supported |

## CustomFunctionsRuntime 1.1, 1.2, and 1.3

The CustomFunctionsRuntime 1.1 is the first version of the API. Requirement set 1.2 adds the `CustomFunctions.Error` object to support error handling. Requirement set 1.3 adds [XLL streaming](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf#custom-function-behavior-for-xll-compatible-functions) support and new `ErrorCode` options to the [CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) object.

## See also

- [Custom Functions Reference Documentation](/javascript/api/custom-functions-runtime)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
