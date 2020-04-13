---
title: DocumentWindow object (PowerPoint)
keywords: vbapp10.chm511000
f1_keywords:
- vbapp10.chm511000
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow
ms.assetid: 567c5e66-8d68-a868-4072-b5358cf69546
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentWindow object (PowerPoint)

Represents a document window. The **DocumentWindow** object is a member of the **[DocumentWindows](PowerPoint.DocumentWindows.md)** collection. The **DocumentWindows** collection contains all the open document windows.


## Remarks

Use the  **[Presentation](PowerPoint.Application.Presentations.md)** property to return the presentation that's currently running in the specified document window.

Use the  **[Selection](PowerPoint.DocumentWindow.Selection.md)** property to return the selection.

Use the  **[SplitHorizontal](PowerPoint.DocumentWindow.SplitHorizontal.md)** property to return the percentage of the screen width that the outline pane occupies in normal view.

Use the  **[SplitVertical](PowerPoint.DocumentWindow.SplitVertical.md)** property to return the percentage of the screen height that the slide pane occupies in normal view.

Use the  **[View](PowerPoint.DocumentWindow.View.md)** property to return the view in the specified document window.


## Example

Use  **Windows** (_index_), where _index_ is the document window index number, to return a single **DocumentWindow** object. The following example activates document window two.


```vb
Windows(2).Activate
```

The first member of the  **DocumentWindows** collection, `Windows(1)`, always returns the active document window. Alternatively, you can use the  **[ActiveWindow](PowerPoint.Application.ActiveWindow.md)** property to return the active document window. The following example maximizes the active window.




```vb
ActiveWindow.WindowState = ppWindowMaximized
```

Use  **Panes** (_index_), where _index_ is the pane index number, to manipulate panes within normal, slide, outline, or notes page views of the document window. The following example activates pane three, which is the notes pane.




```vb
ActiveWindow.Panes(3).Activate
```

Use the  **[ActivePane](PowerPoint.DocumentWindow.ActivePane.md)** property to return the active pane within the document window. The following example checks to see if the active pane is the outline pane. If not, it activates the outline pane.




```vb
mypane = ActiveWindow.ActivePane.ViewType

    If mypane <> 1 Then

        ActiveWindow.Panes(1).Activate

    End If
```


## Methods



|Name|
|:-----|
|**[Activate](PowerPoint.DocumentWindow.Activate.md)**|
|**[Close](PowerPoint.DocumentWindow.Close.md)**|
|**[ExpandSection](PowerPoint.DocumentWindow.ExpandSection.md)**|
|**[FitToPage](PowerPoint.DocumentWindow.FitToPage.md)**|
|**[IsSectionExpanded](PowerPoint.DocumentWindow.IsSectionExpanded.md)**|
|**[LargeScroll](PowerPoint.DocumentWindow.LargeScroll.md)**|
|**[NewWindow](PowerPoint.DocumentWindow.NewWindow.md)**|
|**[PointsToScreenPixelsX](PowerPoint.DocumentWindow.PointsToScreenPixelsX.md)**|
|**[PointsToScreenPixelsY](PowerPoint.DocumentWindow.PointsToScreenPixelsY.md)**|
|**[RangeFromPoint](PowerPoint.DocumentWindow.RangeFromPoint.md)**|
|**[ScrollIntoView](PowerPoint.DocumentWindow.ScrollIntoView.md)**|
|**[SmallScroll](PowerPoint.DocumentWindow.SmallScroll.md)**|

## Properties



|Name|
|:-----|
|**[Active](PowerPoint.DocumentWindow.Active.md)**|
|**[ActivePane](PowerPoint.DocumentWindow.ActivePane.md)**|
|**[Application](PowerPoint.DocumentWindow.Application.md)**|
|**[BlackAndWhite](PowerPoint.DocumentWindow.BlackAndWhite.md)**|
|**[Caption](PowerPoint.DocumentWindow.Caption.md)**|
|**[Height](PowerPoint.DocumentWindow.Height.md)**|
|**[Left](PowerPoint.DocumentWindow.Left.md)**|
|**[Panes](PowerPoint.DocumentWindow.Panes.md)**|
|**[Parent](PowerPoint.DocumentWindow.Parent.md)**|
|**[Presentation](PowerPoint.DocumentWindow.Presentation.md)**|
|**[Selection](PowerPoint.DocumentWindow.Selection.md)**|
|**[SplitHorizontal](PowerPoint.DocumentWindow.SplitHorizontal.md)**|
|**[SplitVertical](PowerPoint.DocumentWindow.SplitVertical.md)**|
|**[Top](PowerPoint.DocumentWindow.Top.md)**|
|**[View](PowerPoint.DocumentWindow.View.md)**|
|**[ViewType](PowerPoint.DocumentWindow.ViewType.md)**|
|**[Width](PowerPoint.DocumentWindow.Width.md)**|
|**[WindowState](PowerPoint.DocumentWindow.WindowState.md)**|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]