---
title: Shape.Id property (PowerPoint)
keywords: vbapp10.chm547069
f1_keywords:
- vbapp10.chm547069
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Id
ms.assetid: b8d1c2ed-08e6-2a1d-7603-d80387fa4ee4
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.Id property (PowerPoint)

Returns a **Long** that identifies the shape or range of shapes. Read-only.


## Syntax

_expression_. `Id`

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Return value

Long


## Example

This example adds a new shape to the active presentation, then fills the shape according to the value of the  **Id** property.


```vb
Sub ShapeID()

    With ActivePresentation.Slides(1).Shapes.AddShape _
            (Type:=msoShape5pointStar, Left:=100, _
            Top:=100, Width:=100, Height:=100)

        Select Case .Id
            Case 0 To 500
                .Fill.ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=0)

            Case 500 To 1000
                .Fill.ForeColor.RGB = RGB(Red:=255, Green:=255, Blue:=0)

            Case 1000 To 1500
                .Fill.ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=255)

            Case 1500 To 2000
                .Fill.ForeColor.RGB = RGB(Red:=0, Green:=255, Blue:=0)

            Case 2000 To 2500
                .Fill.ForeColor.RGB = RGB(Red:=0, Green:=255, Blue:=255)

            Case Else
                .Fill.ForeColor.RGB = RGB(Red:=0, Green:=0, Blue:=255)

        End Select
    End With

End Sub
```


## See also


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]