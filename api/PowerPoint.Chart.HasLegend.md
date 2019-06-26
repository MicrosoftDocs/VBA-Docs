---
title: Chart.HasLegend property (PowerPoint)
keywords: vbapp10.chm684032
f1_keywords:
- vbapp10.chm684032
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.HasLegend
ms.assetid: 084f7de3-b0ed-d7b3-3b24-465e74afa167
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.HasLegend property (PowerPoint)

 **True** if the chart has a legend. Read/write **Boolean**.


## Syntax

_expression_.**HasLegend**

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