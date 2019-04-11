---
title: ChartType property (Excel Graph)
keywords: vbagr10.chm66936
f1_keywords:
- vbagr10.chm66936
ms.prod: excel
api_name:
- Excel.ChartType
ms.assetid: a59871a9-d2f9-657a-1553-eba8c4e4a5a8
ms.date: 04/10/2019
localization_priority: Normal
---


# ChartType property (Excel Graph)

Returns or sets the chart type. Read/write **[XlChartType](excel.xlcharttype.md)**.

## Syntax

_expression_.**ChartType**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the bubble size in chart group one to 200 percent of the default size if the chart is a 2D bubble chart.

```vb
With myChart 
 If .ChartType = xlBubble Then 
 .ChartGroups(1).BubbleScale = 200 
 End If 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
