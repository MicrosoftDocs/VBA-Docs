---
title: Chart.ChartType property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.ChartType
ms.assetid: 5a806b77-1efd-fd3a-132f-f6e3afd7315d
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.ChartType property (PowerPoint)

Returns or sets the chart type. Read/write  **[XlChartType](Excel.XlChartType.md)**.


## Syntax

_expression_.**ChartType**

_expression_ A variable that represents a **[Chart](PowerPoint.Chart.md)** object.


## Remarks

Some chart types are not available for PivotChart reports.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the bubble size in chart group one to 200% of the default size if the chart is a 2D bubble chart.




```vb
With ActiveDocument.InlineShapes(1).Chart 
    If .ChartType = xlBubble Then 
        .ChartGroups(1).BubbleScale = 200 
    End If 
End With
```


## See also


[Chart Object](PowerPoint.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]