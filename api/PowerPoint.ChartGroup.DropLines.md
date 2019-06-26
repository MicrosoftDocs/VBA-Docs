---
title: ChartGroup.DropLines property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.DropLines
ms.assetid: 5646620d-e023-5953-4c91-34234de15b30
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroup.DropLines property (PowerPoint)

Returns the drop lines for a series on a line chart or area chart. Read-only  **[DropLines](PowerPoint.DropLines.md)**.


## Syntax

_expression_.**DropLines**

_expression_ A variable that represents a **[ChartGroup](PowerPoint.ChartGroup.md)** object.


## Remarks

This property applies only to line charts or area charts. 


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables drop lines for chart group one of the first chart in the active document and then sets their line style, weight, and color. You should run the example on a 2D line chart that has one series.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With Chart.ChartGroups(1)

            .HasDropLines = True

            With .DropLines.Border

                .LineStyle = xlThin

                .Weight = xlMedium

                .ColorIndex = 3

            End With

        End With

    End If

End With
```


## See also


[ChartGroup Object](PowerPoint.ChartGroup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]