---
title: Custom Functions requirement sets
description: Details about the Custom Functions requirement sets for Excel JavaScript API.
ms.date: 04/15/2024
ms.topic: overview
ms.localizationpriority: medium
---

# Custom Functions requirement sets

[Custom Functions](/office/dev/add-ins/excel/custom-functions-overview) use separate requirement sets from the core Excel JavaScript APIs. The following table lists the Custom Functions requirement sets, the supported Office client applications, and the **minimum** builds or versions for those applications where applicable.

| Requirement set | Office on the web | Office on Windows<br>(Microsoft 365 subscription) | Office on Windows<br>(retail perpetual) | Office on Windows<br>(volume-licensed perpetual) | Office on Mac | Office on iPad |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| CustomFunctionsRuntime 1.4 | Supported | Version 2208 (Build 15601.20148) | Not supported | Not supported | Version 16.64 (22081401) | Not supported |
| CustomFunctionsRuntime 1.3 | Supported | Version 2008 (Build 13127.20296) | Version 2311 (Build 17029.20126) | Office 2021: Version 2108 (Build 14332.20011) | Version 16.40 (20081000) | Not supported |
| CustomFunctionsRuntime 1.2 | Supported | Version 1909 (Build 11929.20934) | Version 2311 (Build 17029.20126) | Office 2021: Version 2108 (Build 14332.20011) | Version 16.34 (20020900) | Not supported |
| CustomFunctionsRuntime 1.1 | Supported | Version 1903 (Build 11425.20156) | Version 2311 (Build 17029.20126) | Office 2021: Version 2108 (Build 14332.20011) | Version 16.34 (20020900) | Not supported |

## Requirement set history

- Requirement set 1.4 includes [custom functions integration with data types](/office/dev/add-ins/excel/custom-functions-data-types-concepts) and the [`allowCustomDataForDataTypeAny` JSON manifest property](/office/dev/add-ins/excel/custom-functions-json#allowcustomdatafordatatypeany) to support the data types integration.
- Requirement set 1.3 adds [XLL streaming](/office/dev/add-ins/excel/make-custom-functions-compatible-with-xll-udf#custom-function-behavior-for-xll-compatible-functions) support and new `ErrorCode` options to the [CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) object.
- Requirement set 1.2 adds the `CustomFunctions.Error` object to support error handling.
- Requirement set 1.1 is the first version of the API.

## See also

- [Custom Functions Reference Documentation](/javascript/api/custom-functions-runtime)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
