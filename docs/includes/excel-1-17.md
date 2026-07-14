| Class | Fields | Description |
|:---|:---|:---|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[changeRuleToCellValue(properties: Excel.ConditionalCellValueRule)](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletocellvalue-member(1))|Change the conditional format rule type to cell value.|
||[changeRuleToColorScale()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletocolorscale-member(1))|Change the conditional format rule type to color scale.|
||[changeRuleToContainsText(properties: Excel.ConditionalTextComparisonRule)](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletocontainstext-member(1))|Change the conditional format rule type to text comparison.|
||[changeRuleToCustom(formula: string)](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletocustom-member(1))|Change the conditional format rule type to custom.|
||[changeRuleToDataBar()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletodatabar-member(1))|Change the conditional format rule type to data bar.|
||[changeRuleToIconSet()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletoiconset-member(1))|Change the conditional format rule type to icon set.|
||[changeRuleToPresetCriteria(properties: Excel.ConditionalPresetCriteriaRule)](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletopresetcriteria-member(1))|Change the conditional format rule type to preset criteria.|
||[changeRuleToTopBottom(properties: Excel.ConditionalTopBottomRule)](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-changeruletotopbottom-member(1))|Change the conditional format rule type to top/bottom.|
||[setRanges(ranges: Range \| RangeAreas \| string)](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-setranges-member(1))|Set the ranges that the conditional format rule is applied to.|
|[ConditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|[clearFormat()](/javascript/api/excel/excel.conditionalrangeformat#excel-excel-conditionalrangeformat-clearformat-member(1))|Remove the format properties from a conditional format rule.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[currencySymbol](/javascript/api/excel/excel.numberformatinfo#excel-excel-numberformatinfo-currencysymbol-member)|Gets the currency symbol for currency values.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onNameChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onnamechanged-member)|Occurs when the worksheet name is changed.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onvisibilitychanged-member)|Occurs when the worksheet visibility is changed.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onMoved](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onmoved-member)|Occurs when a worksheet is moved within a workbook.|
||[onNameChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onnamechanged-member)|Occurs when the worksheet name is changed in the worksheet collection.|
||[onVisibilityChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onvisibilitychanged-member)|Occurs when the worksheet visibility is changed in the worksheet collection.|
|[WorksheetMovedEventArgs](/javascript/api/excel/excel.worksheetmovedeventargs)|[positionAfter](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-positionafter-member)|Gets the new position of the worksheet, after the move.|
||[positionBefore](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-positionbefore-member)|Gets the previous position of the worksheet, prior to the move.|
||[source](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-source-member)|The source of the event.|
||[type](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetmovedeventargs#excel-excel-worksheetmovedeventargs-worksheetid-member)|Gets the ID of the worksheet that was moved.|
|[WorksheetNameChangedEventArgs](/javascript/api/excel/excel.worksheetnamechangedeventargs)|[nameAfter](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-nameafter-member)|Gets the new name of the worksheet, after the name change.|
||[nameBefore](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-namebefore-member)|Gets the previous name of the worksheet, before the name changed.|
||[source](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-source-member)|The source of the event.|
||[type](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/javascript/api/excel/excel.worksheetnamechangedeventargs#excel-excel-worksheetnamechangedeventargs-worksheetid-member)|Gets the ID of the worksheet with the new name.|
|[WorksheetVisibilityChangedEventArgs](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs)|[source](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-source-member)|The source of the event.|
||[type](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-type-member)|Gets the type of the event.|
||[visibilityAfter](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-visibilityafter-member)|Gets the new visibility setting of the worksheet, after the visibility change.|
||[visibilityBefore](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-visibilitybefore-member)|Gets the previous visibility setting of the worksheet, before the visibility change.|
||[worksheetId](/javascript/api/excel/excel.worksheetvisibilitychangedeventargs#excel-excel-worksheetvisibilitychangedeventargs-worksheetid-member)|Gets the ID of the worksheet whose visibility has changed.|
