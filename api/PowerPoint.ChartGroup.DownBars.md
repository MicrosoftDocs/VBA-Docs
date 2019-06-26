---
title: ChartGroup.DownBars property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.DownBars
ms.assetid: d3af23f5-f408-5f9d-f86d-40ba2c15c870
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroup.DownBars property (PowerPoint)

Returns the down bars on a line chart. Read-only  **[DownBars](PowerPoint.DownBars.md)**.


## Syntax

_expression_.**DownBars**

_expression_ A variable that represents a **[ChartGroup](PowerPoint.ChartGroup.md)** object.


## Remarks

This property applies only to line charts. 


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables up bars and down bars for chart group one of the first chart in the active document and then sets their colors. You should run the example on a 2D line chart that has two series that cross each other at one or more data points.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With Chart.ChartGroups(1)

            .HasUpDownBars = True

            .DownBars.Interior.ColorIndex = 3

            .UpBars.Interior.ColorIndex = 5

        End With

    End If

End With
```


## See also


[ChartGroup Object](PowerPoint.ChartGroup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]