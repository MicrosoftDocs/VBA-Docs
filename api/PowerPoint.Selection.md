---
title: Selection object (PowerPoint)
keywords: vbapp10.chm508000
f1_keywords:
- vbapp10.chm508000
ms.prod: powerpoint
api_name:
- PowerPoint.Selection
ms.assetid: a7def3bd-9dff-da53-152d-4fd686642413
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection object (PowerPoint)

Represents the selection in the specified document window. The **Selection** object is deleted whenever you change slides in an active slide view (the **Type** property will return **ppSelectionNone**).


## Example

Use the [Selection](PowerPoint.Cell.Selected.md)property to return the  **Selection** object. The following example places a copy of the selection in the active window on the Clipboard.


```vb
ActiveWindow.Selection.Copy
```

Use the [ShapeRange](PowerPoint.Selection.ShapeRange.md), [SlideRange](PowerPoint.Selection.SlideRange.md), or [TextRange](PowerPoint.Selection.TextRange.md)property to return a range of shapes, slides, or text from the selection.

The following example sets the fill foreground color for the selected shapes in window two, assuming that there's at least one shape selected, and assuming that all selected shapes have a fill whose forecolor can be set.




```vb
With Windows(2).Selection.ShapeRange.Fill

    .Visible = True

    .ForeColor.RGB = RGB(255, 0, 255)

End With
```

The following example sets the text in the first selected shape in window two if that shape contains a text frame.




```vb
With Windows(2).Selection.ShapeRange(1)

    If .HasTextFrame Then

        .TextFrame.TextRange = "Current Choice"

    End If

End With
```

The following example cuts the selected text in the active window and places it on the Clipboard.




```vb
ActiveWindow.Selection.TextRange.Cut
```

The following example duplicates all the slides in the selection (if you are in slide view, this duplicates the current slide).




```vb
ActiveWindow.Selection.SlideRange.Duplicate
```

If you don't have an object of the appropriate type selected when you use one of these properties (for instance, if you use the  **ShapeRange** property when there are no shapes selected), an error occurs. Use the [Type](PowerPoint.Selection.Type.md)property to determine what kind of object or objects are selected. The following example checks to see whether the selection contains slides. If the selection does contain slides, the example sets the background for the first slide in the selection.




```vb
With Windows(2).Selection

    If .Type = ppSelectionSlides Then

        With .SlideRange(1)

            .FollowMasterBackground = False

            .Background.Fill.PresetGradient _

                msoGradientHorizontal, 1, msoGradientLateSunset

        End With

    End If

End With
```


## Methods



|Name|
|:-----|
|[Copy](PowerPoint.Selection.Copy.md)|
|[Cut](PowerPoint.Selection.Cut.md)|
|[Delete](PowerPoint.Selection.Delete.md)|
|[Unselect](PowerPoint.Selection.Unselect.md)|

## Properties



|Name|
|:-----|
|[Application](PowerPoint.Selection.Application.md)|
|[ChildShapeRange](PowerPoint.Selection.ChildShapeRange.md)|
|[HasChildShapeRange](PowerPoint.Selection.HasChildShapeRange.md)|
|[Parent](PowerPoint.Selection.Parent.md)|
|[ShapeRange](PowerPoint.Selection.ShapeRange.md)|
|[SlideRange](PowerPoint.Selection.SlideRange.md)|
|[TextRange](PowerPoint.Selection.TextRange.md)|
|[TextRange2](PowerPoint.Selection.TextRange2.md)|
|[Type](PowerPoint.Selection.Type.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
