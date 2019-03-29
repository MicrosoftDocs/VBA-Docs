---
title: Pane object (Excel)
keywords: vbaxl10.chm359072
f1_keywords:
- vbaxl10.chm359072
ms.prod: excel
api_name:
- Excel.Pane
ms.assetid: 9064bb89-d08c-bbd3-3c0f-77a39586bbbb
ms.date: 03/30/2019
localization_priority: Normal
---


# Pane object (Excel)

Represents a pane of a window.


## Remarks

**Pane** objects exist only for worksheets and Microsoft Excel 4.0 macro sheets. The **Pane** object is a member of the **[Panes](Excel.Panes.md)** collection. The **Panes** collection contains all of the panes shown in a single window.


## Example

Use **[Panes](Excel.Window.Panes.md)** (_index_), where _index_ is the pane index number, to return a single **Pane** object.

The following example splits the window in which worksheet one is displayed and then scrolls through the pane in the lower-left corner until row five is at the top of the pane.

```vb
Worksheets(1).Activate 
ActiveWindow.Split = True 
ActiveWindow.Panes(3).ScrollRow = 5
```

## Methods

- [Activate](Excel.Pane.Activate.md)
- [LargeScroll](Excel.Pane.LargeScroll.md)
- [PointsToScreenPixelsX](Excel.Pane.PointsToScreenPixelsX.md)
- [PointsToScreenPixelsY](Excel.Pane.PointsToScreenPixelsY.md)
- [ScrollIntoView](Excel.Pane.ScrollIntoView.md)
- [SmallScroll](Excel.Pane.SmallScroll.md)

## Properties

- [Application](Excel.Pane.Application.md)
- [Creator](Excel.Pane.Creator.md)
- [Index](Excel.Pane.Index.md)
- [Parent](Excel.Pane.Parent.md)
- [ScrollColumn](Excel.Pane.ScrollColumn.md)
- [ScrollRow](Excel.Pane.ScrollRow.md)
- [VisibleRange](Excel.Pane.VisibleRange.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]