---
title: Window Object (Excel)
keywords: vbaxl10.chm355072
f1_keywords:
- vbaxl10.chm355072
ms.prod: excel
api_name:
- Excel.Window
ms.assetid: 8591b1ad-76f8-14e2-9120-406b65093f5a
ms.date: 06/08/2017
---


# Window Object (Excel)

Represents a window.


## Remarks

 Many worksheet characteristics, such as scroll bars and gridlines, are actually properties of the window. The **Window** object is a member of the **[Windows](Excel.Windows.md)** collection. The **Windows** collection for the **Application** object contains all the windows in the application, whereas the **Windows** collection for the **Workbook** object contains only the windows in the specified workbook.


## Example

Use  **Windows** ( _index_ ), where _index_ is the window name or index number, to return a single **Window** object. The following example maximizes the active window.


```
Windows(1).WindowState = xlMaximized
```

Note that the active window is always  `Windows(1)`.

The window caption is the text shown in the title bar at the top of the window when the window isn't maximized. The caption is also shown in the list of open files on the bottom of the  **Windows** menu. Use the **[Caption](Excel.Window.Caption.md)** property to set or return the window caption. Changing the window caption doesn't change the name of the workbook. The following example turns off cell gridlines for the worksheet shown in the Book1.xls:1 window.




```
Windows("book1.xls":1).DisplayGridlines = False
```


## Methods



|**Name**|
|:-----|
|[Activate](Excel.Window.Activate.md)|
|[ActivateNext](Excel.Window.ActivateNext.md)|
|[ActivatePrevious](Excel.Window.ActivatePrevious.md)|
|[Close](Excel.Window.Close.md)|
|[LargeScroll](Excel.Window.LargeScroll.md)|
|[NewWindow](Excel.Window.NewWindow.md)|
|[PointsToScreenPixelsX](Excel.Window.PointsToScreenPixelsX.md)|
|[PointsToScreenPixelsY](Excel.Window.PointsToScreenPixelsY.md)|
|[PrintOut](Excel.Window.PrintOut.md)|
|[PrintPreview](Excel.Window.PrintPreview.md)|
|[RangeFromPoint](Excel.Window.RangeFromPoint.md)|
|[ScrollIntoView](Excel.Window.ScrollIntoView.md)|
|[ScrollWorkbookTabs](Excel.Window.ScrollWorkbookTabs.md)|
|[SmallScroll](Excel.Window.SmallScroll.md)|

## Properties



|**Name**|
|:-----|
|[ActiveCell](Excel.Window.ActiveCell.md)|
|[ActiveChart](Excel.Window.ActiveChart.md)|
|[ActivePane](Excel.Window.ActivePane.md)|
|[ActiveSheet](Excel.Window.ActiveSheet.md)|
|[ActiveSheetView](Excel.Window.ActiveSheetView.md)|
|[Application](Excel.Window.Application.md)|
|[AutoFilterDateGrouping](Excel.Window.AutoFilterDateGrouping.md)|
|[Caption](Excel.Window.Caption.md)|
|[Creator](Excel.Window.Creator.md)|
|[DisplayFormulas](Excel.Window.DisplayFormulas.md)|
|[DisplayGridlines](Excel.Window.DisplayGridlines.md)|
|[DisplayHeadings](Excel.Window.DisplayHeadings.md)|
|[DisplayHorizontalScrollBar](Excel.Window.DisplayHorizontalScrollBar.md)|
|[DisplayOutline](Excel.Window.DisplayOutline.md)|
|[DisplayRightToLeft](Excel.Window.DisplayRightToLeft.md)|
|[DisplayRuler](Excel.Window.DisplayRuler.md)|
|[DisplayVerticalScrollBar](Excel.Window.DisplayVerticalScrollBar.md)|
|[DisplayWhitespace](Excel.Window.DisplayWhitespace.md)|
|[DisplayWorkbookTabs](Excel.Window.DisplayWorkbookTabs.md)|
|[DisplayZeros](Excel.Window.DisplayZeros.md)|
|[EnableResize](Excel.Window.EnableResize.md)|
|[FreezePanes](Excel.Window.FreezePanes.md)|
|[GridlineColor](Excel.Window.GridlineColor.md)|
|[GridlineColorIndex](Excel.Window.GridlineColorIndex.md)|
|[Height](Excel.Window.Height.md)|
|[Hwnd](Excel.window.hwnd.md)|
|[Index](Excel.Window.Index.md)|
|[Left](Excel.Window.Left.md)|
|[OnWindow](Excel.Window.OnWindow.md)|
|[Panes](Excel.Window.Panes.md)|
|[Parent](Excel.Window.Parent.md)|
|[RangeSelection](Excel.Window.RangeSelection.md)|
|[ScrollColumn](Excel.Window.ScrollColumn.md)|
|[ScrollRow](Excel.Window.ScrollRow.md)|
|[SelectedSheets](Excel.Window.SelectedSheets.md)|
|[Selection](Excel.Window.Selection.md)|
|[SheetViews](Excel.Window.SheetViews.md)|
|[Split](Excel.Window.Split.md)|
|[SplitColumn](Excel.Window.SplitColumn.md)|
|[SplitHorizontal](Excel.Window.SplitHorizontal.md)|
|[SplitRow](Excel.Window.SplitRow.md)|
|[SplitVertical](Excel.Window.SplitVertical.md)|
|[TabRatio](Excel.Window.TabRatio.md)|
|[Top](Excel.Window.Top.md)|
|[Type](Excel.Window.Type.md)|
|[UsableHeight](Excel.Window.UsableHeight.md)|
|[UsableWidth](Excel.Window.UsableWidth.md)|
|[View](Excel.Window.View.md)|
|[Visible](Excel.Window.Visible.md)|
|[VisibleRange](Excel.Window.VisibleRange.md)|
|[Width](Excel.Window.Width.md)|
|[WindowNumber](Excel.Window.WindowNumber.md)|
|[WindowState](Excel.Window.WindowState.md)|
|[Zoom](window-zoom-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
