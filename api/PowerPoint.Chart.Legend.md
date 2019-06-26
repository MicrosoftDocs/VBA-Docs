---
title: Chart.Legend property (PowerPoint)
keywords: vbapp10.chm684035
f1_keywords:
- vbapp10.chm684035
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.Legend
ms.assetid: 1bd67a75-9dd4-2d8c-99b5-82bc91cf85d9
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.Legend property (PowerPoint)

Returns the legend for the chart. Read-only  **[Legend](PowerPoint.Legend.md)**.


## Syntax

_expression_.**Legend**

_expression_ A variable that represents a **[Chart](PowerPoint.Chart.md)** object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables the legend for the first chart in the active document and then sets the legend font color to blue.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart

            .HasLegend = True

            .Legend.Font.ColorIndex = 5

        End With

    End If

End With
```


## See also


[Chart Object](PowerPoint.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]