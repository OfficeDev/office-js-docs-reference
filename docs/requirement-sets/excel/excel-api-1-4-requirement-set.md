---
title: Excel JavaScript API requirement set 1.4
description: Details about the ExcelApi 1.4 requirement set.
ms.date: 11/09/2020
ms.topic: whats-new
ms.prod: excel
ms.localizationpriority: medium
---

# What's new in Excel JavaScript API 1.4

The following are the new additions to the Excel JavaScript APIs in requirement set 1.4.

## Named item add and new properties

New properties:

* `comment`
* `scope` - Worksheet or workbook scoped items.
* `worksheet` - Returns the worksheet on which the named item is scoped to.

New methods:

* `add(name: string, reference: Range or string, comment: string)` - Adds a new name to the collection of the given scope.
* `addFormulaLocal(name: string, formula: string, comment: string)` - Adds a new name to the collection of the given scope using the user's locale for the formula.

## Settings API in the Excel namespace

The [Setting](/javascript/api/excel/excel.setting) object represents a key:value pair for a setting persisted to the document. The functionality of `Excel.Setting` is equivalent to `Office.Settings`, but uses the batched API syntax, rather than the Common API's callback model.

APIs include `getItem()` to get setting entry via the key and `add()` to add the specified key:value setting pair to the workbook.

## Others

* Set the table column name.
* Add a table column to the end of the table.
* Add multiple rows to a table at a time.
* `range.getColumnsAfter(count: number)` and `range.getColumnsBefore(count: number)` to get a certain number of columns to the right/left of the current Range object.
* The [\*OrNullObject methods and properties](/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties): This functionality allows getting an object using a key. If the object does not exist, the returned object's `isNullObject` property will be true. This allows developers to check if an object exists without having to handle it through exception handling. An `*OrNullObject` method is available on most collection objects.

```js
worksheet.getItemOrNullObject("itemName")
```

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.4. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.4 or earlier, see [Excel APIs in requirement set 1.4 or earlier](/javascript/api/excel?view=excel-js-1.4&preserve-view=true).

[!INCLUDE[API table](../../includes/excel-1_4.md)]

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.4&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
