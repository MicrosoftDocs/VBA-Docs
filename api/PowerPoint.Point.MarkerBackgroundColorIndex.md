---
title: Point.MarkerBackgroundColorIndex Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Point.MarkerBackgroundColorIndex
ms.assetid: 357a97f9-d20a-6761-5977-23ee526a277a
ms.date: 06/08/2017
localization_priority: Normal
---


# Point.MarkerBackgroundColorIndex Property (PowerPoint)

Returns or sets the marker background color as an index into the current color palette, or as one of the following  **[xlColorIndex](PowerPoint.XlColorIndex.md)** constants: **xlColorIndexAutomatic** or **xlColorIndexNone**. Read/write **Long**.


## Syntax

 _expression_. `MarkerBackgroundColorIndex`

 _expression_ A variable that represents a '[Point](PowerPoint.Point.md)' object.


## Remarks

The property applies only to line, scatter, and radar charts. 


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the marker background and foreground colors for the second point in series one for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.SeriesCollection(1).Points(2)

            ' Set the background color to green.

            .MarkerBackgroundColorIndex = 4



            ' Set the foreground color to red.

            .MarkerForegroundColorIndex = 3

        End With

    End If

End With


```


## See also


[Point Object](PowerPoint.Point.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]