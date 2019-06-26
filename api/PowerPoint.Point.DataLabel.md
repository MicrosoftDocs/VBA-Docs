---
title: Point.DataLabel property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Point.DataLabel
ms.assetid: 0f202f4c-2627-09e0-38d8-fd51aa1cdfb1
ms.date: 06/08/2017
localization_priority: Normal
---


# Point.DataLabel property (PowerPoint)

Returns the data label associated with the point. Read-only  **[DataLabel](PowerPoint.DataLabel.md)**.


## Syntax

_expression_.**DataLabel**

_expression_ A variable that represents a '[Point](PowerPoint.Point.md)' object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables the data label for point seven in series three of the first chart in the active document, and then it sets the data label color to blue.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.SeriesCollection(3).Points(7)

            .HasDataLabel = True

            .ApplyDataLabels type:=xlValue

            .DataLabel.Font.ColorIndex = 5

        End With

    End If

End With
```


## See also


[Point Object](PowerPoint.Point.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]