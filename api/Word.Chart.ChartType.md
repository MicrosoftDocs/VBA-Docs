---
title: Chart.ChartType property (Word)
keywords: vbawd10.chm79365496
f1_keywords:
- vbawd10.chm79365496
ms.prod: word
api_name:
- Word.Chart.ChartType
ms.assetid: ad75b5bc-b323-8f67-cf1a-b4d6b6969eed
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.ChartType property (Word)

Returns or sets the chart type. Read/write  **[XlChartType](Excel.XlChartType.md)**.


## Syntax

_expression_.**ChartType**

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Remarks

Some chart types are not available for PivotChart reports.


## Example

The following example sets the bubble size in chart group one to 200% of the default size if the chart is a 2D bubble chart.


```vb
With ActiveDocument.InlineShapes(1).Chart 
 If .ChartType = xlBubble Then 
 .ChartGroups(1).BubbleScale = 200 
 End If 
End With
```


## See also


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]