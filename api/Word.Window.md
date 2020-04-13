---
title: Window object (Word)
keywords: vbawd10.chm2402
f1_keywords:
- vbawd10.chm2402
ms.prod: word
api_name:
- Word.Window
ms.assetid: d92f83f9-ae44-56c0-4584-7a9359253c6d
ms.date: 06/08/2017
localization_priority: Normal
---


# Window object (Word)

Represents a window. Many document characteristics, such as scroll bars and rulers, are actually properties of the window.


## Remarks

The **Window** object is a member of the **[Windows](Word.windows.md)** collection. The **Windows** collection for the **Application** object contains all the windows in the application, whereas the **Windows** collection for the **Document** object contains only the windows that display the specified document.

Use  **Windows** (Index), where Index is the window name or the index number, to return a single **Window** object. The following example maximizes the Document1 window.




```vb
Windows("Document1").WindowState = wdWindowStateMaximize
```

The index number is the number to the left of the window name on the **Window** menu. The following example displays the caption of the first window in the **Windows** collection.




```vb
MsgBox Windows(1).Caption
```

Use the **Add** method or the **NewWindow** method to add a new window to the **Windows** collection. Each of the following statements creates a new window for the document in the active window.




```vb
ActiveDocument.ActiveWindow.NewWindow 
NewWindow 
Windows.Add
```

A colon (:) and a number appear in the window caption when more than one window is open for a document.

When you switch the view to print preview, a new window is created. This window is removed from the **Windows** collection when you close print preview.


## Methods



|Name|
|:-----|
|[Activate](Word.Window.Activate.md)|
|[Close](Word.Window.Close.md)|
|[GetPoint](Word.Window.GetPoint.md)|
|[LargeScroll](Word.Window.LargeScroll.md)|
|[NewWindow](Word.Window.NewWindow.md)|
|[PageScroll](Word.Window.PageScroll.md)|
|[PrintOut](Word.Window.PrintOut.md)|
|[RangeFromPoint](Word.Window.RangeFromPoint.md)|
|[ScrollIntoView](Word.Window.ScrollIntoView.md)|
|[SetFocus](Word.Window.SetFocus.md)|
|[SmallScroll](Word.Window.SmallScroll.md)|
|[ToggleRibbon](Word.Window.ToggleRibbon.md)|

## Properties



|Name|
|:-----|
|[Active](Word.Window.Active.md)|
|[ActivePane](Word.Window.ActivePane.md)|
|[Application](Word.Window.Application.md)|
|[Caption](Word.Window.Caption.md)|
|[Creator](Word.Window.Creator.md)|
|[DisplayHorizontalScrollBar](Word.Window.DisplayHorizontalScrollBar.md)|
|[DisplayLeftScrollBar](Word.Window.DisplayLeftScrollBar.md)|
|[DisplayRightRuler](Word.Window.DisplayRightRuler.md)|
|[DisplayRulers](Word.Window.DisplayRulers.md)|
|[DisplayScreenTips](Word.Window.DisplayScreenTips.md)|
|[DisplayVerticalRuler](Word.Window.DisplayVerticalRuler.md)|
|[DisplayVerticalScrollBar](Word.Window.DisplayVerticalScrollBar.md)|
|[Document](Word.Window.Document.md)|
|[DocumentMap](Word.Window.DocumentMap.md)|
|[EnvelopeVisible](Word.Window.EnvelopeVisible.md)|
|[Height](Word.Window.Height.md)|
|[HorizontalPercentScrolled](Word.Window.HorizontalPercentScrolled.md)|
|[Hwnd](Word.window.hwnd.md)|
|[IMEMode](Word.Window.IMEMode.md)|
|[Index](Word.Window.Index.md)|
|[Left](Word.Window.Left.md)|
|[Next](Word.Window.Next.md)|
|[Panes](Word.Window.Panes.md)|
|[Parent](Word.Window.Parent.md)|
|[Previous](Word.Window.Previous.md)|
|[Selection](Word.Window.Selection.md)|
|[ShowSourceDocuments](Word.Window.ShowSourceDocuments.md)|
|[Split](Word.Window.Split.md)|
|[SplitVertical](Word.Window.SplitVertical.md)|
|[StyleAreaWidth](Word.Window.StyleAreaWidth.md)|
|[Thumbnails](Word.Window.Thumbnails.md)|
|[Top](Word.Window.Top.md)|
|[Type](Word.Window.Type.md)|
|[UsableHeight](Word.Window.UsableHeight.md)|
|[UsableWidth](Word.Window.UsableWidth.md)|
|[VerticalPercentScrolled](Word.Window.VerticalPercentScrolled.md)|
|[View](Word.Window.View.md)|
|[Visible](Word.Window.Visible.md)|
|[Width](Word.Window.Width.md)|
|[WindowNumber](Word.WindowNumber.md)|
|[WindowState](Word.Window.WindowState.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]