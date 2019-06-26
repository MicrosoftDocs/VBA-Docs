---
title: Legend.Position property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Legend.Position
ms.assetid: 82d71eda-aa17-b463-9934-6f79fe370f67
ms.date: 06/08/2017
localization_priority: Normal
---


# Legend.Position property (PowerPoint)

Returns or sets the position of the legend on the chart. Read/write  **[XlLegendPosition](PowerPoint.XlLegendPosition.md)**.


## Syntax

_expression_.**Position**

_expression_ A variable that represents a '[Legend](PowerPoint.Legend.md)' object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example moves the chart legend to the bottom of the chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Legend.Position = xlLegendPositionBottom

    End If

End With
```


## See also


[Legend Object](PowerPoint.Legend.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]