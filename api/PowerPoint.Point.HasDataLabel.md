---
title: Point.HasDataLabel property (PowerPoint)
keywords: vbapp10.chm65613
f1_keywords:
- vbapp10.chm65613
ms.prod: powerpoint
api_name:
- PowerPoint.Point.HasDataLabel
ms.assetid: bb0e96e7-5280-c074-5278-f8e5acb7bab3
ms.date: 06/08/2017
localization_priority: Normal
---


# Point.HasDataLabel property (PowerPoint)

 **True** if the point has a data label. Read/write **Boolean**.


## Syntax

_expression_.**HasDataLabel**

_expression_ A variable that represents a '[Point](PowerPoint.Point.md)' object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables the data label for point seven in series three for the first chart in the active document, and then it sets the data label color to blue.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With Chart.SeriesCollection(3).Points(7)

            .HasDataLabel = True

            .ApplyDataLabels Type:=xlValue

            .DataLabel.Font.ColorIndex = 5

        End With

    End If

End With
```


## See also


[Point Object](PowerPoint.Point.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]