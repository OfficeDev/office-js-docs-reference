| Class | Fields | Description |
|:---|:---|:---|
|[Application](/.application)|[suspendApiCalculationUntilNextSync()](/.application#excel-javascript/api/excel/-application-suspendapicalculationuntilnextsync-member(1))|Suspends calculation until the next `context.sync()` is called.|
|[CellValueConditionalFormat](/.cellvalueconditionalformat)|[format](/.cellvalueconditionalformat#excel-javascript/api/excel/-cellvalueconditionalformat-format-member)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/.cellvalueconditionalformat#excel-javascript/api/excel/-cellvalueconditionalformat-rule-member)|Specifies the rule object on this conditional format.|
|[ColorScaleConditionalFormat](/.colorscaleconditionalformat)|[criteria](/.colorscaleconditionalformat#excel-javascript/api/excel/-colorscaleconditionalformat-criteria-member)|The criteria of the color scale.|
||[threeColorScale](/.colorscaleconditionalformat#excel-javascript/api/excel/-colorscaleconditionalformat-threecolorscale-member)|If `true`, the color scale will have three points (minimum, midpoint, maximum), otherwise it will have two (minimum, maximum).|
|[ConditionalCellValueRule](/.conditionalcellvaluerule)|[formula1](/.conditionalcellvaluerule#excel-javascript/api/excel/-conditionalcellvaluerule-formula1-member)|The formula, if required, on which to evaluate the conditional format rule.|
||[formula2](/.conditionalcellvaluerule#excel-javascript/api/excel/-conditionalcellvaluerule-formula2-member)|The formula, if required, on which to evaluate the conditional format rule.|
||[operator](/.conditionalcellvaluerule#excel-javascript/api/excel/-conditionalcellvaluerule-operator-member)|The operator of the cell value conditional format.|
|[ConditionalColorScaleCriteria](/.conditionalcolorscalecriteria)|[maximum](/.conditionalcolorscalecriteria#excel-javascript/api/excel/-conditionalcolorscalecriteria-maximum-member)|The maximum point of the color scale criterion.|
||[midpoint](/.conditionalcolorscalecriteria#excel-javascript/api/excel/-conditionalcolorscalecriteria-midpoint-member)|The midpoint of the color scale criterion, if the color scale is a 3-color scale.|
||[minimum](/.conditionalcolorscalecriteria#excel-javascript/api/excel/-conditionalcolorscalecriteria-minimum-member)|The minimum point of the color scale criterion.|
|[ConditionalColorScaleCriterion](/.conditionalcolorscalecriterion)|[color](/.conditionalcolorscalecriterion#excel-javascript/api/excel/-conditionalcolorscalecriterion-color-member)|HTML color code representation of the color scale color (e.g., #FF0000 represents Red).|
||[formula](/.conditionalcolorscalecriterion#excel-javascript/api/excel/-conditionalcolorscalecriterion-formula-member)|A number, a formula, or `null` (if `type` is `lowestValue`).|
||[type](/.conditionalcolorscalecriterion#excel-javascript/api/excel/-conditionalcolorscalecriterion-type-member)|What the criterion conditional formula should be based on.|
|[ConditionalDataBarNegativeFormat](/.conditionaldatabarnegativeformat)|[borderColor](/.conditionaldatabarnegativeformat#excel-javascript/api/excel/-conditionaldatabarnegativeformat-bordercolor-member)|HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[fillColor](/.conditionaldatabarnegativeformat#excel-javascript/api/excel/-conditionaldatabarnegativeformat-fillcolor-member)|HTML color code representing the fill color, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[matchPositiveBorderColor](/.conditionaldatabarnegativeformat#excel-javascript/api/excel/-conditionaldatabarnegativeformat-matchpositivebordercolor-member)|Specifies if the negative data bar has the same border color as the positive data bar.|
||[matchPositiveFillColor](/.conditionaldatabarnegativeformat#excel-javascript/api/excel/-conditionaldatabarnegativeformat-matchpositivefillcolor-member)|Specifies if the negative data bar has the same fill color as the positive data bar.|
|[ConditionalDataBarPositiveFormat](/.conditionaldatabarpositiveformat)|[borderColor](/.conditionaldatabarpositiveformat#excel-javascript/api/excel/-conditionaldatabarpositiveformat-bordercolor-member)|HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[fillColor](/.conditionaldatabarpositiveformat#excel-javascript/api/excel/-conditionaldatabarpositiveformat-fillcolor-member)|HTML color code representing the fill color, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[gradientFill](/.conditionaldatabarpositiveformat#excel-javascript/api/excel/-conditionaldatabarpositiveformat-gradientfill-member)|Specifies if the data bar has a gradient.|
|[ConditionalDataBarRule](/.conditionaldatabarrule)|[formula](/.conditionaldatabarrule#excel-javascript/api/excel/-conditionaldatabarrule-formula-member)|The formula, if required, on which to evaluate the data bar rule.|
||[type](/.conditionaldatabarrule#excel-javascript/api/excel/-conditionaldatabarrule-type-member)|The type of rule for the data bar.|
|[ConditionalFormat](/.conditionalformat)|[cellValue](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-cellvalue-member)|Returns the cell value conditional format properties if the current conditional format is a `CellValue` type.|
||[cellValueOrNullObject](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-cellvalueornullobject-member)|Returns the cell value conditional format properties if the current conditional format is a `CellValue` type.|
||[colorScale](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-colorscale-member)|Returns the color scale conditional format properties if the current conditional format is a `ColorScale` type.|
||[colorScaleOrNullObject](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-colorscaleornullobject-member)|Returns the color scale conditional format properties if the current conditional format is a `ColorScale` type.|
||[custom](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-custom-member)|Returns the custom conditional format properties if the current conditional format is a custom type.|
||[customOrNullObject](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-customornullobject-member)|Returns the custom conditional format properties if the current conditional format is a custom type.|
||[dataBar](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-databar-member)|Returns the data bar properties if the current conditional format is a data bar.|
||[dataBarOrNullObject](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-databarornullobject-member)|Returns the data bar properties if the current conditional format is a data bar.|
||[delete()](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-delete-member(1))|Deletes this conditional format.|
||[getRange()](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-getrange-member(1))|Returns the range the conditional format is applied to.|
||[getRangeOrNullObject()](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-getrangeornullobject-member(1))|Returns the range to which the conditional format is applied.|
||[iconSet](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-iconset-member)|Returns the icon set conditional format properties if the current conditional format is an `IconSet` type.|
||[iconSetOrNullObject](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-iconsetornullobject-member)|Returns the icon set conditional format properties if the current conditional format is an `IconSet` type.|
||[id](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-id-member)|The priority of the conditional format in the current `ConditionalFormatCollection`.|
||[preset](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-preset-member)|Returns the preset criteria conditional format.|
||[presetOrNullObject](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-presetornullobject-member)|Returns the preset criteria conditional format.|
||[priority](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-priority-member)|The priority (or index) within the conditional format collection that this conditional format currently exists in.|
||[stopIfTrue](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-stopiftrue-member)|If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.|
||[textComparison](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-textcomparison-member)|Returns the specific text conditional format properties if the current conditional format is a text type.|
||[textComparisonOrNullObject](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-textcomparisonornullobject-member)|Returns the specific text conditional format properties if the current conditional format is a text type.|
||[topBottom](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-topbottom-member)|Returns the top/bottom conditional format properties if the current conditional format is a `TopBottom` type.|
||[topBottomOrNullObject](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-topbottomornullobject-member)|Returns the top/bottom conditional format properties if the current conditional format is a `TopBottom` type.|
||[type](/.conditionalformat#excel-javascript/api/excel/-conditionalformat-type-member)|A type of conditional format.|
|[ConditionalFormatCollection](/.conditionalformatcollection)|[add(type: Excel.ConditionalFormatType)](/.conditionalformatcollection#excel-javascript/api/excel/-conditionalformatcollection-add-member(1))|Adds a new conditional format to the collection at the first/top priority.|
||[clearAll()](/.conditionalformatcollection#excel-javascript/api/excel/-conditionalformatcollection-clearall-member(1))|Clears all conditional formats active on the current specified range.|
||[getCount()](/.conditionalformatcollection#excel-javascript/api/excel/-conditionalformatcollection-getcount-member(1))|Returns the number of conditional formats in the workbook.|
||[getItem(id: string)](/.conditionalformatcollection#excel-javascript/api/excel/-conditionalformatcollection-getitem-member(1))|Returns a conditional format for the given ID.|
||[getItemAt(index: number)](/.conditionalformatcollection#excel-javascript/api/excel/-conditionalformatcollection-getitemat-member(1))|Returns a conditional format at the given index.|
||[items](/.conditionalformatcollection#excel-javascript/api/excel/-conditionalformatcollection-items-member)|Gets the loaded child items in this collection.|
|[ConditionalFormatRule](/.conditionalformatrule)|[formula](/.conditionalformatrule#excel-javascript/api/excel/-conditionalformatrule-formula-member)|The formula, if required, on which to evaluate the conditional format rule.|
||[formulaLocal](/.conditionalformatrule#excel-javascript/api/excel/-conditionalformatrule-formulalocal-member)|The formula, if required, on which to evaluate the conditional format rule in the user's language.|
||[formulaR1C1](/.conditionalformatrule#excel-javascript/api/excel/-conditionalformatrule-formular1c1-member)|The formula, if required, on which to evaluate the conditional format rule in R1C1-style notation.|
|[ConditionalIconCriterion](/.conditionaliconcriterion)|[customIcon](/.conditionaliconcriterion#excel-javascript/api/excel/-conditionaliconcriterion-customicon-member)|The custom icon for the current criterion, if different from the default icon set, else `null` will be returned.|
||[formula](/.conditionaliconcriterion#excel-javascript/api/excel/-conditionaliconcriterion-formula-member)|A number or a formula depending on the type.|
||[operator](/.conditionaliconcriterion#excel-javascript/api/excel/-conditionaliconcriterion-operator-member)|`greaterThan` or `greaterThanOrEqual` for each of the rule types for the icon conditional format.|
||[type](/.conditionaliconcriterion#excel-javascript/api/excel/-conditionaliconcriterion-type-member)|What the icon conditional formula should be based on.|
|[ConditionalPresetCriteriaRule](/.conditionalpresetcriteriarule)|[criterion](/.conditionalpresetcriteriarule#excel-javascript/api/excel/-conditionalpresetcriteriarule-criterion-member)|The criterion of the conditional format.|
|[ConditionalRangeBorder](/.conditionalrangeborder)|[color](/.conditionalrangeborder#excel-javascript/api/excel/-conditionalrangeborder-color-member)|HTML color code representing the color of the border line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[sideIndex](/.conditionalrangeborder#excel-javascript/api/excel/-conditionalrangeborder-sideindex-member)|Constant value that indicates the specific side of the border.|
||[style](/.conditionalrangeborder#excel-javascript/api/excel/-conditionalrangeborder-style-member)|One of the constants of line style specifying the line style for the border.|
|[ConditionalRangeBorderCollection](/.conditionalrangebordercollection)|[bottom](/.conditionalrangebordercollection#excel-javascript/api/excel/-conditionalrangebordercollection-bottom-member)|Gets the bottom border.|
||[count](/.conditionalrangebordercollection#excel-javascript/api/excel/-conditionalrangebordercollection-count-member)|Number of border objects in the collection.|
||[getItem(index: Excel.ConditionalRangeBorderIndex)](/.conditionalrangebordercollection#excel-javascript/api/excel/-conditionalrangebordercollection-getitem-member(1))|Gets a border object using its name.|
||[getItemAt(index: number)](/.conditionalrangebordercollection#excel-javascript/api/excel/-conditionalrangebordercollection-getitemat-member(1))|Gets a border object using its index.|
||[items](/.conditionalrangebordercollection#excel-javascript/api/excel/-conditionalrangebordercollection-items-member)|Gets the loaded child items in this collection.|
||[left](/.conditionalrangebordercollection#excel-javascript/api/excel/-conditionalrangebordercollection-left-member)|Gets the left border.|
||[right](/.conditionalrangebordercollection#excel-javascript/api/excel/-conditionalrangebordercollection-right-member)|Gets the right border.|
||[top](/.conditionalrangebordercollection#excel-javascript/api/excel/-conditionalrangebordercollection-top-member)|Gets the top border.|
|[ConditionalRangeFill](/.conditionalrangefill)|[clear()](/.conditionalrangefill#excel-javascript/api/excel/-conditionalrangefill-clear-member(1))|Resets the fill.|
||[color](/.conditionalrangefill#excel-javascript/api/excel/-conditionalrangefill-color-member)|HTML color code representing the color of the fill, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
|[ConditionalRangeFont](/.conditionalrangefont)|[bold](/.conditionalrangefont#excel-javascript/api/excel/-conditionalrangefont-bold-member)|Specifies if the font is bold.|
||[clear()](/.conditionalrangefont#excel-javascript/api/excel/-conditionalrangefont-clear-member(1))|Resets the font formats.|
||[color](/.conditionalrangefont#excel-javascript/api/excel/-conditionalrangefont-color-member)|HTML color code representation of the text color (e.g., #FF0000 represents Red).|
||[italic](/.conditionalrangefont#excel-javascript/api/excel/-conditionalrangefont-italic-member)|Specifies if the font is italic.|
||[strikethrough](/.conditionalrangefont#excel-javascript/api/excel/-conditionalrangefont-strikethrough-member)|Specifies the strikethrough status of the font.|
||[underline](/.conditionalrangefont#excel-javascript/api/excel/-conditionalrangefont-underline-member)|The type of underline applied to the font.|
|[ConditionalRangeFormat](/.conditionalrangeformat)|[borders](/.conditionalrangeformat#excel-javascript/api/excel/-conditionalrangeformat-borders-member)|Collection of border objects that apply to the overall conditional format range.|
||[fill](/.conditionalrangeformat#excel-javascript/api/excel/-conditionalrangeformat-fill-member)|Returns the fill object defined on the overall conditional format range.|
||[font](/.conditionalrangeformat#excel-javascript/api/excel/-conditionalrangeformat-font-member)|Returns the font object defined on the overall conditional format range.|
||[numberFormat](/.conditionalrangeformat#excel-javascript/api/excel/-conditionalrangeformat-numberformat-member)|Represents Excel's number format code for the given range.|
|[ConditionalTextComparisonRule](/.conditionaltextcomparisonrule)|[operator](/.conditionaltextcomparisonrule#excel-javascript/api/excel/-conditionaltextcomparisonrule-operator-member)|The operator of the text conditional format.|
||[text](/.conditionaltextcomparisonrule#excel-javascript/api/excel/-conditionaltextcomparisonrule-text-member)|The text value of the conditional format.|
|[ConditionalTopBottomRule](/.conditionaltopbottomrule)|[rank](/.conditionaltopbottomrule#excel-javascript/api/excel/-conditionaltopbottomrule-rank-member)|The rank between 1 and 1000 for numeric ranks or 1 and 100 for percent ranks.|
||[type](/.conditionaltopbottomrule#excel-javascript/api/excel/-conditionaltopbottomrule-type-member)|Format values based on the top or bottom rank.|
|[CustomConditionalFormat](/.customconditionalformat)|[format](/.customconditionalformat#excel-javascript/api/excel/-customconditionalformat-format-member)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/.customconditionalformat#excel-javascript/api/excel/-customconditionalformat-rule-member)|Specifies the `Rule` object on this conditional format.|
|[DataBarConditionalFormat](/.databarconditionalformat)|[axisColor](/.databarconditionalformat#excel-javascript/api/excel/-databarconditionalformat-axiscolor-member)|HTML color code representing the color of the Axis line, in the form #RRGGBB (e.g., "FFA500") or as a named HTML color (e.g., "orange").|
||[axisFormat](/.databarconditionalformat#excel-javascript/api/excel/-databarconditionalformat-axisformat-member)|Representation of how the axis is determined for an Excel data bar.|
||[barDirection](/.databarconditionalformat#excel-javascript/api/excel/-databarconditionalformat-bardirection-member)|Specifies the direction that the data bar graphic should be based on.|
||[lowerBoundRule](/.databarconditionalformat#excel-javascript/api/excel/-databarconditionalformat-lowerboundrule-member)|The rule for what constitutes the lower bound (and how to calculate it, if applicable) for a data bar.|
||[negativeFormat](/.databarconditionalformat#excel-javascript/api/excel/-databarconditionalformat-negativeformat-member)|Representation of all values to the left of the axis in an Excel data bar.|
||[positiveFormat](/.databarconditionalformat#excel-javascript/api/excel/-databarconditionalformat-positiveformat-member)|Representation of all values to the right of the axis in an Excel data bar.|
||[showDataBarOnly](/.databarconditionalformat#excel-javascript/api/excel/-databarconditionalformat-showdatabaronly-member)|If `true`, hides the values from the cells where the data bar is applied.|
||[upperBoundRule](/.databarconditionalformat#excel-javascript/api/excel/-databarconditionalformat-upperboundrule-member)|The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.|
|[IconSetConditionalFormat](/.iconsetconditionalformat)|[criteria](/.iconsetconditionalformat#excel-javascript/api/excel/-iconsetconditionalformat-criteria-member)|An array of criteria and icon sets for the rules and potential custom icons for conditional icons.|
||[reverseIconOrder](/.iconsetconditionalformat#excel-javascript/api/excel/-iconsetconditionalformat-reverseiconorder-member)|If `true`, reverses the icon orders for the icon set.|
||[showIconOnly](/.iconsetconditionalformat#excel-javascript/api/excel/-iconsetconditionalformat-showicononly-member)|If `true`, hides the values and only shows icons.|
||[style](/.iconsetconditionalformat#excel-javascript/api/excel/-iconsetconditionalformat-style-member)|If set, displays the icon set option for the conditional format.|
|[PresetCriteriaConditionalFormat](/.presetcriteriaconditionalformat)|[format](/.presetcriteriaconditionalformat#excel-javascript/api/excel/-presetcriteriaconditionalformat-format-member)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties.|
||[rule](/.presetcriteriaconditionalformat#excel-javascript/api/excel/-presetcriteriaconditionalformat-rule-member)|The rule of the conditional format.|
|[Range](/.range)|[calculate()](/.range#excel-javascript/api/excel/-range-calculate-member(1))|Calculates a range of cells on a worksheet.|
||[conditionalFormats](/.range#excel-javascript/api/excel/-range-conditionalformats-member)|The collection of `ConditionalFormats` that intersect the range.|
|[TextConditionalFormat](/.textconditionalformat)|[format](/.textconditionalformat#excel-javascript/api/excel/-textconditionalformat-format-member)|Returns a format object, encapsulating the conditional format's font, fill, borders, and other properties.|
||[rule](/.textconditionalformat#excel-javascript/api/excel/-textconditionalformat-rule-member)|The rule of the conditional format.|
|[TopBottomConditionalFormat](/.topbottomconditionalformat)|[format](/.topbottomconditionalformat#excel-javascript/api/excel/-topbottomconditionalformat-format-member)|Returns a format object, encapsulating the conditional format's font, fill, borders, and other properties.|
||[rule](/.topbottomconditionalformat#excel-javascript/api/excel/-topbottomconditionalformat-rule-member)|The criteria of the top/bottom conditional format.|
|[Worksheet](/.worksheet)|[calculate(markAllDirty: boolean)](/.worksheet#excel-javascript/api/excel/-worksheet-calculate-member(1))|Calculates all cells on a worksheet.|
