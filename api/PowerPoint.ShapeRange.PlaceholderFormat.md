---
title: ShapeRange.PlaceholderFormat property (PowerPoint)
keywords: vbapp10.chm548046
f1_keywords:
- vbapp10.chm548046
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.PlaceholderFormat
ms.assetid: 3c3c344f-aa02-29b2-5ef5-d090f3e32a2c
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.PlaceholderFormat property (PowerPoint)

Returns a **[PlaceholderFormat](PowerPoint.PlaceholderFormat.md)** object that contains the properties that are unique to placeholders. Read-only.


## Syntax

_expression_. `PlaceholderFormat`

_expression_ A variable that represents a **[ShapeRange](PowerPoint.ShapeRange.md)** object.


## Return value

PlaceholderFormat


## Example

This example adds text to placeholder one on slide one in the active presentation if that placeholder is a horizontal title placeholder.


```vb
With ActivePresentation.Slides(1).Shapes.Placeholders

    If .Count > 0 Then
        With .Item(1)
            Select Case .PlaceholderFormat.Type

                Case ppPlaceholderTitle
                    .TextFrame.TextRange = "Title Text"

                Case ppPlaceholderCenterTitle
                    .TextFrame.TextRange = "Centered Title Text"

                Case Else
                    MsgBox "There's no horizontal" & _
                        "title on this slide"

            End Select
        End With
    End If

End With
```


## See also


[ShapeRange Object](PowerPoint.ShapeRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]