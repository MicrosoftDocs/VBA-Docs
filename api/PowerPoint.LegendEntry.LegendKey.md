---
title: LegendEntry.LegendKey property (PowerPoint)
keywords: vbapp10.chm65710
f1_keywords:
- vbapp10.chm65710
ms.prod: powerpoint
api_name:
- PowerPoint.LegendEntry.LegendKey
ms.assetid: 6265569c-fc7c-5fe8-864e-d543a08b33f4
ms.date: 06/08/2017
localization_priority: Normal
---


# LegendEntry.LegendKey property (PowerPoint)

Returns the legend key that is associated with the entry. Read-only  **[LegendKey](PowerPoint.LegendKey.md)**.


## Syntax

_expression_. `LegendKey`

_expression_ A variable that represents a '[LegendEntry](PowerPoint.LegendEntry.md)' object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the legend key for legend entry one on the first chart in the active document to be a triangle. You should run the example on a 2D line chart.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.Legend.LegendEntries(1).LegendKey _
            .MarkerStyle = xlMarkerStyleTriangle
    End If
End With
```


## See also


[LegendEntry Object](PowerPoint.LegendEntry.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]