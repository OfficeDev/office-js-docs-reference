| Class | Fields | Description |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[activeWindow](/javascript/api/excel/excel.application#excel-excel-application-activewindow-member)|Returns a `window` object that represents the active window (the window on top).|
||[checkSpelling(word: string, options?: Excel.CheckSpellingOptions)](/javascript/api/excel/excel.application#excel-excel-application-checkspelling-member(1))|Checks the spelling of a single word.|
||[enterEditingMode()](/javascript/api/excel/excel.application#excel-excel-application-entereditingmode-member(1))|Enters editing mode for the selected range in the active worksheet.|
||[union(firstRange: Range \| RangeAreas, secondRange: Range \| RangeAreas, ...additionalRanges: (Range \| RangeAreas)[])](/javascript/api/excel/excel.application#excel-excel-application-union-member(1))|Returns a `RangeAreas` object that represents the union of two or more `Range` or `RangeAreas` objects.|
||[windows](/javascript/api/excel/excel.application#excel-excel-application-windows-member)|Returns all the open Excel windows.|
|[CheckSpellingOptions](/javascript/api/excel/excel.checkspellingoptions)|[customDictionary](/javascript/api/excel/excel.checkspellingoptions#excel-excel-checkspellingoptions-customdictionary-member)|Optional.|
||[ignoreUppercase](/javascript/api/excel/excel.checkspellingoptions#excel-excel-checkspellingoptions-ignoreuppercase-member)|Optional.|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooterPicture](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-centerfooterpicture-member)|Gets a `HeaderFooterPicture` object that represents the picture for the center section of the footer.|
||[centerHeaderPicture](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-centerheaderpicture-member)|Gets a `HeaderFooterPicture` object that represents the picture for the center section of the header.|
||[leftFooterPicture](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-leftfooterpicture-member)|Gets a `HeaderFooterPicture` object that represents the picture for the left section of the footer.|
||[leftHeaderPicture](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-leftheaderpicture-member)|Gets a `HeaderFooterPicture` object that represents the picture for the left section of the header.|
||[rightFooterPicture](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-rightfooterpicture-member)|Gets a `HeaderFooterPicture` object that represents the picture for the right section of the footer.|
||[rightHeaderPicture](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-rightheaderpicture-member)|Gets a `HeaderFooterPicture` object that represents the picture for the right section of the header.|
|[HeaderFooterPicture](/javascript/api/excel/excel.headerfooterpicture)|[brightness](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-brightness-member)|Specifies the brightness of the picture.|
||[colorType](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-colortype-member)|Specifies the type of color transformation of the picture.|
||[contrast](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-contrast-member)|Specifies the contrast of the picture.|
||[cropBottom](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-cropbottom-member)|Specifies the number of points that are cropped off the bottom of the picture.|
||[cropLeft](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-cropleft-member)|Specifies the number of points that are cropped off the left side of the picture.|
||[cropRight](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-cropright-member)|Specifies the number of points that are cropped off the right side of the picture.|
||[cropTop](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-croptop-member)|Specifies the number of points that are cropped off the top of the picture.|
||[filename](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-filename-member)|Specifies the URL (on the intranet or the web) or path (local or network) to the location where the source object is saved.|
||[height](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-height-member)|Specifies the height of the picture in points.|
||[lockAspectRatio](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-lockaspectratio-member)|Specifies a value that indicates whether the picture retains its original proportions when resized.|
||[width](/javascript/api/excel/excel.headerfooterpicture#excel-excel-headerfooterpicture-width-member)|Specifies the width of the picture in points.|
|[Image](/javascript/api/excel/excel.image)|[brightness](/javascript/api/excel/excel.image#excel-excel-image-brightness-member)|Specifies the brightness of the image.|
||[colorType](/javascript/api/excel/excel.image#excel-excel-image-colortype-member)|Specifies the type of color transformation applied to the image.|
||[contrast](/javascript/api/excel/excel.image#excel-excel-image-contrast-member)|Specifies the contrast of the image.|
||[cropBottom](/javascript/api/excel/excel.image#excel-excel-image-cropbottom-member)|Specifies the number of points that are cropped off the bottom of the image.|
||[cropLeft](/javascript/api/excel/excel.image#excel-excel-image-cropleft-member)|Specifies the number of points that are cropped off the left side of the image.|
||[cropRight](/javascript/api/excel/excel.image#excel-excel-image-cropright-member)|Specifies the number of points that are cropped off the right side of the image.|
||[cropTop](/javascript/api/excel/excel.image#excel-excel-image-croptop-member)|Specifies the number of points that are cropped off the top of the image.|
||[incrementBrightness(increment: number)](/javascript/api/excel/excel.image#excel-excel-image-incrementbrightness-member(1))|Increments the brightness of the image by a specified amount.|
||[incrementContrast(increment: number)](/javascript/api/excel/excel.image#excel-excel-image-incrementcontrast-member(1))|Increments the contrast of the image by a specified amount.|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[alignMarginsHeaderFooter](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-alignmarginsheaderfooter-member)|Specifies whether Excel aligns the header and the footer with the margins set in the page setup options.|
||[printQuality](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printquality-member)|Specifies a two-element array that contains both horizontal and vertical print quality values.|
|[Pane](/javascript/api/excel/excel.pane)|[index](/javascript/api/excel/excel.pane#excel-excel-pane-index-member)|Returns index of the pane.|
|[PaneCollection](/javascript/api/excel/excel.panecollection)|[getCount()](/javascript/api/excel/excel.panecollection#excel-excel-panecollection-getcount-member(1))|Returns the number of panes in the collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.panecollection#excel-excel-panecollection-getitemat-member(1))|Gets the pane in the collection by index.|
||[items](/javascript/api/excel/excel.panecollection#excel-excel-panecollection-items-member)|Gets the loaded child items in this collection.|
|[Range](/javascript/api/excel/excel.range)|[checkSpelling(options?: Excel.CheckSpellingOptions)](/javascript/api/excel/excel.range#excel-excel-range-checkspelling-member(1))|Checks the spelling of words in this range.|
||[formulaArray](/javascript/api/excel/excel.range#excel-excel-range-formulaarray-member)|Specifies the array formula of a range.|
||[showDependents(remove?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-showdependents-member(1))|Draws tracer arrows to the direct dependents of the range.|
||[showPrecedents(remove?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-showprecedents-member(1))|Draws tracer arrows to the direct precedents of the range.|
|[Window](/javascript/api/excel/excel.window)|[activate()](/javascript/api/excel/excel.window#excel-excel-window-activate-member(1))|Activates the window.|
||[activateNext()](/javascript/api/excel/excel.window#excel-excel-window-activatenext-member(1))|Activates the next window.|
||[activatePrevious()](/javascript/api/excel/excel.window#excel-excel-window-activateprevious-member(1))|Activates the previous window.|
||[activeCell](/javascript/api/excel/excel.window#excel-excel-window-activecell-member)|Gets the active cell in the window.|
||[activePane](/javascript/api/excel/excel.window#excel-excel-window-activepane-member)|Gets the active pane in the window.|
||[activeWorksheet](/javascript/api/excel/excel.window#excel-excel-window-activeworksheet-member)|Gets the active worksheet in the window.|
||[autoFilterDateGroupingEnabled](/javascript/api/excel/excel.window#excel-excel-window-autofilterdategroupingenabled-member)|Specifies whether AutoFilter date grouping is enabled in the window.|
||[close()](/javascript/api/excel/excel.window#excel-excel-window-close-member(1))|Closes the window.|
||[enableResize](/javascript/api/excel/excel.window#excel-excel-window-enableresize-member)|Specifies whether resizing is enabled for the window.|
||[freezePanes](/javascript/api/excel/excel.window#excel-excel-window-freezepanes-member)|Specifies whether panes are frozen in the window.|
||[height](/javascript/api/excel/excel.window#excel-excel-window-height-member)|Specifies the height of the window.|
||[index](/javascript/api/excel/excel.window#excel-excel-window-index-member)|Gets the index of the window.|
||[isVisible](/javascript/api/excel/excel.window#excel-excel-window-isvisible-member)|Specifies whether the window is visible.|
||[largeScroll(Down: number, Up: number, ToRight: number, ToLeft: number)](/javascript/api/excel/excel.window#excel-excel-window-largescroll-member(1))|Scrolls the window by multiple pages.|
||[left](/javascript/api/excel/excel.window#excel-excel-window-left-member)|Specifies the distance, in points, from the left edge of the computer screen to the left edge of the window.|
||[name](/javascript/api/excel/excel.window#excel-excel-window-name-member)|Specifies the name of the window.|
||[newWindow()](/javascript/api/excel/excel.window#excel-excel-window-newwindow-member(1))|Opens a new Excel window.|
||[panes](/javascript/api/excel/excel.window#excel-excel-window-panes-member)|Gets a collection of panes associated with the window.|
||[pointsToScreenPixelsX(Points: number)](/javascript/api/excel/excel.window#excel-excel-window-pointstoscreenpixelsx-member(1))|Converts horizontal points to screen pixels.|
||[pointsToScreenPixelsY(Points: number)](/javascript/api/excel/excel.window#excel-excel-window-pointstoscreenpixelsy-member(1))|Converts vertical points to screen pixels.|
||[scrollColumn](/javascript/api/excel/excel.window#excel-excel-window-scrollcolumn-member)|Specifies the scroll column of the window.|
||[scrollIntoView(Left: number, Top: number, Width: number, Height: number, Start?: boolean)](/javascript/api/excel/excel.window#excel-excel-window-scrollintoview-member(1))|Scrolls the window to bring the specified range into view.|
||[scrollRow](/javascript/api/excel/excel.window#excel-excel-window-scrollrow-member)|Specifies the scroll row of the window.|
||[scrollWorkbookTabs(Sheets?: number, Position?: Excel.ScrollWorkbookTabPosition)](/javascript/api/excel/excel.window#excel-excel-window-scrollworkbooktabs-member(1))|Scrolls the workbook tabs.|
||[showFormulas](/javascript/api/excel/excel.window#excel-excel-window-showformulas-member)|Specifies whether formulas are shown in the window.|
||[showGridlines](/javascript/api/excel/excel.window#excel-excel-window-showgridlines-member)|Specifies whether gridlines are shown in the window.|
||[showHeadings](/javascript/api/excel/excel.window#excel-excel-window-showheadings-member)|Specifies whether headings are shown in the window.|
||[showHorizontalScrollBar](/javascript/api/excel/excel.window#excel-excel-window-showhorizontalscrollbar-member)|Specifies whether the horizontal scroll bar is shown in the window.|
||[showOutline](/javascript/api/excel/excel.window#excel-excel-window-showoutline-member)|Specifies whether outline is shown in the window.|
||[showRightToLeft](/javascript/api/excel/excel.window#excel-excel-window-showrighttoleft-member)|Gets the right-to-left layout value of the window.|
||[showRuler](/javascript/api/excel/excel.window#excel-excel-window-showruler-member)|Specifies whether the ruler is shown in the window.|
||[showVerticalScrollBar](/javascript/api/excel/excel.window#excel-excel-window-showverticalscrollbar-member)|Specifies whether the vertical scroll bar is shown in the window.|
||[showWhitespace](/javascript/api/excel/excel.window#excel-excel-window-showwhitespace-member)|Specifies whether whitespace is shown in the window.|
||[showWorkbookTabs](/javascript/api/excel/excel.window#excel-excel-window-showworkbooktabs-member)|Specifies whether workbook tabs are shown in the window.|
||[showZeros](/javascript/api/excel/excel.window#excel-excel-window-showzeros-member)|Specifies whether zeroes are shown in the window.|
||[smallScroll(Down: number, Up: number, ToRight: number, ToLeft: number)](/javascript/api/excel/excel.window#excel-excel-window-smallscroll-member(1))|Scrolls the window by a number of rows or columns.|
||[split](/javascript/api/excel/excel.window#excel-excel-window-split-member)|Specifies the split state of the window.|
||[splitColumn](/javascript/api/excel/excel.window#excel-excel-window-splitcolumn-member)|Specifies the split column of the window.|
||[splitHorizontal](/javascript/api/excel/excel.window#excel-excel-window-splithorizontal-member)|Specifies the horizontal split of the window.|
||[splitRow](/javascript/api/excel/excel.window#excel-excel-window-splitrow-member)|Specifies the split row of the window.|
||[splitVertical](/javascript/api/excel/excel.window#excel-excel-window-splitvertical-member)|Specifies the vertical split of the window.|
||[tabRatio](/javascript/api/excel/excel.window#excel-excel-window-tabratio-member)|Specifies the tab ratio of the window.|
||[top](/javascript/api/excel/excel.window#excel-excel-window-top-member)|Specifies the distance, in points, from the top edge of the window to the top edge of the usable area (below the menus, any toolbars docked at the top, and the formula bar).|
||[type](/javascript/api/excel/excel.window#excel-excel-window-type-member)|Specifies the type of the window.|
||[usableHeight](/javascript/api/excel/excel.window#excel-excel-window-usableheight-member)|Gets the usable height of the window.|
||[usableWidth](/javascript/api/excel/excel.window#excel-excel-window-usablewidth-member)|Gets the usable width of the window.|
||[view](/javascript/api/excel/excel.window#excel-excel-window-view-member)|Specifies the view of the window.|
||[visibleRange](/javascript/api/excel/excel.window#excel-excel-window-visiblerange-member)|Gets the visible range of the window.|
||[width](/javascript/api/excel/excel.window#excel-excel-window-width-member)|Specifies the display width of the window.|
||[windowNumber](/javascript/api/excel/excel.window#excel-excel-window-windownumber-member)|Gets the window number.|
||[windowState](/javascript/api/excel/excel.window#excel-excel-window-windowstate-member)|Specifies the window state.|
||[zoom](/javascript/api/excel/excel.window#excel-excel-window-zoom-member)|Specifies an integer value that represents the display size of the window.|
|[WindowCollection](/javascript/api/excel/excel.windowcollection)|[breakSideBySide()](/javascript/api/excel/excel.windowcollection#excel-excel-windowcollection-breaksidebyside-member(1))|Breaks the side-by-side view of windows.|
||[compareCurrentSideBySideWith(windowName: string)](/javascript/api/excel/excel.windowcollection#excel-excel-windowcollection-comparecurrentsidebysidewith-member(1))|Compares the current window side by side with the specified window.|
||[getCount()](/javascript/api/excel/excel.windowcollection#excel-excel-windowcollection-getcount-member(1))|Gets the number of windows in the collection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.windowcollection#excel-excel-windowcollection-getitemat-member(1))|Gets the Window in the collection by index.|
||[items](/javascript/api/excel/excel.windowcollection#excel-excel-windowcollection-items-member)|Gets the loaded child items in this collection.|
||[resetPositionsSideBySide()](/javascript/api/excel/excel.windowcollection#excel-excel-windowcollection-resetpositionssidebyside-member(1))|Resets the positions of windows in side-by-side view.|
|[Workbook](/javascript/api/excel/excel.workbook)|[focus()](/javascript/api/excel/excel.workbook#excel-excel-workbook-focus-member(1))|Sets focus on the workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[checkSpelling(options?: Excel.CheckSpellingOptions)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-checkspelling-member(1))|Checks the spelling of words in this worksheet.|
||[clearArrows()](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-cleararrows-member(1))|Clears the tracer arrows from the worksheet.|
||[evaluate(name: string)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-evaluate-member(1))|Returns the evaluation result of a formula string.|
