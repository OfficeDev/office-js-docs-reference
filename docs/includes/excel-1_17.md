| Class | Fields | Description |
|:---|:---|:---|
|[ConditionalFormat](/.conditionalformat)|[changeRuleToCellValue(properties: Excel.ConditionalCellValueRule)](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-changeruletocellvalue-member(1))|Change the conditional format rule type to cell value.|
||[changeRuleToColorScale()](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-changeruletocolorscale-member(1))|Change the conditional format rule type to color scale.|
||[changeRuleToContainsText(properties: Excel.ConditionalTextComparisonRule)](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-changeruletocontainstext-member(1))|Change the conditional format rule type to text comparison.|
||[changeRuleToCustom(formula: string)](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-changeruletocustom-member(1))|Change the conditional format rule type to custom.|
||[changeRuleToDataBar()](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-changeruletodatabar-member(1))|Change the conditional format rule type to data bar.|
||[changeRuleToIconSet()](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-changeruletoiconset-member(1))|Change the conditional format rule type to icon set.|
||[changeRuleToPresetCriteria(properties: Excel.ConditionalPresetCriteriaRule)](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-changeruletopresetcriteria-member(1))|Change the conditional format rule type to preset criteria.|
||[changeRuleToTopBottom(properties: Excel.ConditionalTopBottomRule)](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-changeruletotopbottom-member(1))|Change the conditional format rule type to top/bottom.|
||[setRanges(ranges: Range \| RangeAreas \| string)](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-setranges-member(1))|Set the ranges that the conditional format rule is applied to.|
|[ConditionalRangeFormat](/.conditionalrangeformat)|[clearFormat()](/.conditionalrangeformat#excel-javascript/api/excel/-conditionalrangeformat-clearformat-member(1))|Remove the format properties from a conditional format rule.|
|[NumberFormatInfo](/.numberformatinfo)|[currencySymbol](/.numberformatinfo#excel-javascript/api/excel/-numberformatinfo-currencysymbol-member)|Gets the currency symbol for currency values.|
|[Worksheet](/.worksheet)|[onNameChanged](/.worksheet#excel-javascript/api/excel/-worksheet-onnamechanged-member)|Occurs when the worksheet name is changed.|
||[onVisibilityChanged](/.worksheet#excel-javascript/api/excel/-worksheet-onvisibilitychanged-member)|Occurs when the worksheet visibility is changed.|
|[WorksheetCollection](/.worksheetcollection)|[onMoved](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-onmoved-member)|Occurs when a worksheet is moved within a workbook.|
||[onNameChanged](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-onnamechanged-member)|Occurs when the worksheet name is changed in the worksheet collection.|
||[onVisibilityChanged](/.worksheetcollection#excel-javascript/api/excel/-worksheetcollection-onvisibilitychanged-member)|Occurs when the worksheet visibility is changed in the worksheet collection.|
|[WorksheetMovedEventArgs](/.worksheetmovedeventargs)|[positionAfter](/.worksheetmovedeventargs#excel-javascript/api/excel/-worksheetmovedeventargs-positionafter-member)|Gets the new position of the worksheet, after the move.|
||[positionBefore](/.worksheetmovedeventargs#excel-javascript/api/excel/-worksheetmovedeventargs-positionbefore-member)|Gets the previous position of the worksheet, prior to the move.|
||[source](/.worksheetmovedeventargs#excel-javascript/api/excel/-worksheetmovedeventargs-source-member)|The source of the event.|
||[type](/.worksheetmovedeventargs#excel-javascript/api/excel/-worksheetmovedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.worksheetmovedeventargs#excel-javascript/api/excel/-worksheetmovedeventargs-worksheetid-member)|Gets the ID of the worksheet that was moved.|
|[WorksheetNameChangedEventArgs](/.worksheetnamechangedeventargs)|[nameAfter](/.worksheetnamechangedeventargs#excel-javascript/api/excel/-worksheetnamechangedeventargs-nameafter-member)|Gets the new name of the worksheet, after the name change.|
||[nameBefore](/.worksheetnamechangedeventargs#excel-javascript/api/excel/-worksheetnamechangedeventargs-namebefore-member)|Gets the previous name of the worksheet, before the name changed.|
||[source](/.worksheetnamechangedeventargs#excel-javascript/api/excel/-worksheetnamechangedeventargs-source-member)|The source of the event.|
||[type](/.worksheetnamechangedeventargs#excel-javascript/api/excel/-worksheetnamechangedeventargs-type-member)|Gets the type of the event.|
||[worksheetId](/.worksheetnamechangedeventargs#excel-javascript/api/excel/-worksheetnamechangedeventargs-worksheetid-member)|Gets the ID of the worksheet with the new name.|
|[WorksheetVisibilityChangedEventArgs](/.worksheetvisibilitychangedeventargs)|[source](/.worksheetvisibilitychangedeventargs#excel-javascript/api/excel/-worksheetvisibilitychangedeventargs-source-member)|The source of the event.|
||[type](/.worksheetvisibilitychangedeventargs#excel-javascript/api/excel/-worksheetvisibilitychangedeventargs-type-member)|Gets the type of the event.|
||[visibilityAfter](/.worksheetvisibilitychangedeventargs#excel-javascript/api/excel/-worksheetvisibilitychangedeventargs-visibilityafter-member)|Gets the new visibility setting of the worksheet, after the visibility change.|
||[visibilityBefore](/.worksheetvisibilitychangedeventargs#excel-javascript/api/excel/-worksheetvisibilitychangedeventargs-visibilitybefore-member)|Gets the previous visibility setting of the worksheet, before the visibility change.|
||[worksheetId](/.worksheetvisibilitychangedeventargs#excel-javascript/api/excel/-worksheetvisibilitychangedeventargs-worksheetid-member)|Gets the ID of the worksheet whose visibility has changed.|
