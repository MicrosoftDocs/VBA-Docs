---
title: TextFrame object (PowerPoint)
keywords: vbapp10.chm558000
f1_keywords:
- vbapp10.chm558000
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame
ms.assetid: 03346e81-71b2-0b9e-843d-fb8aa0e3c868
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame object (PowerPoint)

Represents the text frame in a  **Shape** object. Contains the text in the text frame and the properties and methods that control the alignment and anchoring of the text frame.


## Example

Use the  **TextFrame** property to return a **TextFrame** object. The following example adds a rectangle to myDocument, adds text to the rectangle, and then sets the margins for the text frame.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes _

        .AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame

    .TextRange.Text = "Here is some test text"

    .MarginBottom = 10

    .MarginLeft = 10

    .MarginRight = 10

    .MarginTop = 10

End With
```

Use the [HasTextFrame](PowerPoint.Shape.HasTextFrame.md)property to determine whether a shape has a text frame, and use the [HasText](PowerPoint.TextFrame.HasText.md)property to determine whether the text frame contains text, as shown in the following example.




```vb
Set myDocument = ActivePresentation.Slides(1)

For Each s In myDocument.Shapes

    If s.HasTextFrame Then

        With s.TextFrame

            If .HasText Then MsgBox .TextRange.Text

        End With

    End If

Next
```


## Methods



|Name|
|:-----|
|[DeleteText](PowerPoint.TextFrame.DeleteText.md)|

## Properties



|Name|
|:-----|
|[Application](PowerPoint.TextFrame.Application.md)|
|[AutoSize](PowerPoint.TextFrame.AutoSize.md)|
|[Creator](PowerPoint.TextFrame.Creator.md)|
|[HasText](PowerPoint.TextFrame.HasText.md)|
|[HorizontalAnchor](PowerPoint.TextFrame.HorizontalAnchor.md)|
|[MarginBottom](PowerPoint.TextFrame.MarginBottom.md)|
|[MarginLeft](PowerPoint.TextFrame.MarginLeft.md)|
|[MarginRight](PowerPoint.TextFrame.MarginRight.md)|
|[MarginTop](PowerPoint.TextFrame.MarginTop.md)|
|[Orientation](PowerPoint.TextFrame.Orientation.md)|
|[Parent](PowerPoint.TextFrame.Parent.md)|
|[Ruler](PowerPoint.TextFrame.Ruler.md)|
|[TextRange](PowerPoint.TextFrame.TextRange.md)|
|[VerticalAnchor](PowerPoint.TextFrame.VerticalAnchor.md)|
|[WordWrap](PowerPoint.TextFrame.WordWrap.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
