---
title: Excel JavaScript API requirement set 1.16
description: Details about the ExcelApi 1.16 requirement set.
ms.date: 10/10/2022
ms.prod: excel
ms.localizationpriority: medium
---

# What's new in Excel JavaScript API 1.16

The ExcelApi 1.16 added the [data types APIs](/office/dev/add-ins/excel/excel-data-types-overview). With the data types APIs, Excel cells can contain [images from the web](/office/dev/add-ins/excel/excel-data-types-concepts#web-image-values), [formatted number values](/office/dev/add-ins/excel/excel-data-types-concepts#formatted-number-values) that retain their format throughout calculations, and most notably, entity cards. Entity cards extend the potential of Excel add-ins beyond a 2-dimensional grid. They display an icon within a cell that opens a card modal window in the Excel UI when selected. To learn more, see [Use cards with entity value data types](/office/dev/add-ins/excel/excel-data-types-entity-card).

The following table provides a concise summary of the APIs, while the subsequent [API list](#api-list) table gives a detailed list.

| Feature area | Description | Relevant objects |
|:--- |:--- |:--- |
| [Data types](/office/dev/add-ins/excel/excel-data-types-overview) | An extension of existing Excel data types, including support for formatted numbers and web images. | [ArrayCellValue](/javascript/api/excel/excel.arraycellvalue), [BooleanCellValue](/javascript/api/excel/excel.booleancellvalue), [CellValueAttributionAttributes](/javascript/api/excel/excel.cellvalueattributionattributes), [CellValueProviderAttributes](/javascript/api/excel/excel.cellvalueproviderattributes), [DoubleCellValue](/javascript/api/excel/excel.doublecellvalue), [EmptyCellValue](/javascript/api/excel/excel.emptycellvalue), [EntityCellValue](/javascript/api/excel/excel.entitycellvalue), [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue), [RootReferenceCellValue](/javascript/api/excel/excel.rootreferencecellvalue), [StringCellValue](/javascript/api/excel/excel.stringcellvalue), [ValueTypeNotAvailableCellValue](/javascript/api/excel/excel.valuetypenotavailablecellvalue), [WebImageCellValue](/javascript/api/excel/excel.webimagecellvalue) |
| [Data types errors](/office/dev/add-ins/excel/excel-data-types-concepts#improved-error-support) | Error objects that support expanded data types. | [BlockedErrorCellValue](/javascript/api/excel/excel.blockederrorcellvalue), [BusyErrorCellValue](/javascript/api/excel/excel.busyerrorcellvalue), [CalcErrorCellValue](/javascript/api/excel/excel.calcerrorcellvalue), [ConnectErrorCellValue](/javascript/api/excel/excel.connecterrorcellvalue), [Div0ErrorCellValue](/javascript/api/excel/excel.div0errorcellvalue), [FieldErrorCellValue](/javascript/api/excel/excel.fielderrorcellvalue), [GettingDataErrorCellValue](/javascript/api/excel/excel.gettingdataerrorcellvalue), [NotAvailableErrorCellValue](/javascript/api/excel/excel.notavailableerrorcellvalue), [NameErrorCellValue](/javascript/api/excel/excel.nameerrorcellvalue), [NullErrorCellValue](/javascript/api/excel/excel.nullerrorcellvalue), [NumErrorCellValue](/javascript/api/excel/excel.numerrorcellvalue), [PlaceholderErrorCellValue](/javascript/api/excel/excel.placeholdererrorcellvalue), [RefErrorCellValue](/javascript/api/excel/excel.referrorcellvalue), [SpillErrorCellValue](/javascript/api/excel/excel.spillerrorcellvalue), [ValueErrorCellValue](/javascript/api/excel/excel.valueerrorcellvalue)|
| [Entity value data types](/office/dev/add-ins/excel/excel-data-types-concepts#entity-values) | An entity value is a container for data types. Card layout objects manage the display of entity values. | [EntityCellValue](/javascript/api/excel/excel.entitycellvalue), [EntityCardLayout](/javascript/api/excel/excel.entitycardlayout), [EntityPropertyExtraProperties](/javascript/api/excel/excel.entitypropertyextraproperties), [EntityViewLayouts](/javascript/api/excel/excel.entityviewlayouts), [CardLayoutListSection](/javascript/api/excel/excel.cardlayoutlistsection), [CardLayoutPropertyReference](/javascript/api/excel/excel.cardlayoutpropertyreference), [CardLayoutSectionStandardProperties](/javascript/api/excel/excel.cardlayoutsectionstandardproperties), [CardLayoutStandardProperties](/javascript/api/excel/excel.cardlayoutstandardproperties), [CardLayoutTableSection](/javascript/api/excel/excel.cardlayouttablesection) |
| Linked data types | Adds support for data types connected to Excel from external sources. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype), [LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs), [LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection) |

## API list

The following table lists the APIs in Excel JavaScript API requirement set 1.16. To view API reference documentation for all APIs supported by Excel JavaScript API requirement set 1.16 or earlier, see [Excel APIs in requirement set 1.16 or earlier](/javascript/api/excel?view=excel-js-1.16&preserve-view=true).

[!INCLUDE[API table](../../includes/excel-1_16.md)]

## See also

- [Excel JavaScript API Reference Documentation](/javascript/api/excel?view=excel-js-1.16&preserve-view=true)
- [Excel JavaScript API requirement sets](excel-api-requirement-sets.md)
