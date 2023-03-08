---
title: Excel JavaScript API requirement set 1.5
description: Details about the ExcelApi 1.5 requirement set.
ms.date: 03/19/2021
ms.topic: whats-new
ms.prod: excel
ms.localizationpriority: medium
---

# What's new in Excel JavaScript API 1.5

ExcelApi 1.5 adds Custom XML parts. These are accessible through the [custom XML parts collection](/javascript/api/excel/excel.workbook#excel-excel-workbook-customxmlparts-member) in the workbook object.

## Custom XML part

* Get custom XML parts using their ID.
* Get a new scoped collection of custom XML parts whose namespaces match the given namespace.
* Get an XML string associated with a part.
* Provide the ID and namespace of a part.
* Add a new custom XML part to the workbook.
* Set an entire XML part.
* Delete a custom XML part.
* Delete an attribute with the given name from the element identified by xpath.
* Query the XML content by xpath.
* Insert, update, and delete attributes.

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.5. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.5 or earlier, see [Excel APIs in requirement set 1.5 or earlier](/javascript/api/excel?view=excel-js-1.5&preserve-view=true).

[!INCLUDE[API table](../../includes/excel-1_5.md)]

## See also

* [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.5&preserve-view=true)
* [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
