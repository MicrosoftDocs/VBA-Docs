---
title: ChartGroup.SeriesLines property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.SeriesLines
ms.assetid: 40282a82-5912-b5b1-b556-a53c66483502
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroup.SeriesLines property (PowerPoint)

Returns the series lines for a 2D stacked bar, 2D stacked column, pie-of-pie, or bar-of-pie chart. Read-only  **[SeriesLines](PowerPoint.SeriesLines.md)**.


## Syntax

_expression_.**SeriesLines**

_expression_ A variable that represents a **[ChartGroup](PowerPoint.ChartGroup.md)** object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables series lines for chart group one of the first chart in the active document, and then sets the line style, weight, and color of the series lines. You should run the example on a 2D stacked column chart that has two or more series.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.ChartGroups(1)

            .HasSeriesLines = True

            With .SeriesLines.Border

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