---
title: ChartGroup.SeriesCollection method (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.SeriesCollection
ms.assetid: 5d20f5b2-cd4c-06b6-a49c-0ab331157b2f
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroup.SeriesCollection method (PowerPoint)

Returns all the series in the chart group.


## Syntax

_expression_.**SeriesCollection** (_Index_)

_expression_ A variable that represents a **[ChartGroup](PowerPoint.ChartGroup.md)** object.


## Return value

A  **[SeriesCollection](PowerPoint.SeriesCollection.md)** object that represents all the series in the chart group.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example turns on data labels for the first series of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.ChartGroups(1). _
            SeriesCollection(1).HasDataLabels = True
    End If
End With
```


## See also


[ChartGroup Object](PowerPoint.ChartGroup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]