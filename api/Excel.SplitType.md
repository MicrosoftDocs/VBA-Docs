---
title: SplitType property (Excel Graph)
keywords: vbagr10.chm3077588
f1_keywords:
- vbagr10.chm3077588
ms.prod: excel
api_name:
- Excel.SplitType
ms.assetid: e6af8aac-bd1f-9e00-abd7-54e49623d536
ms.date: 04/12/2019
localization_priority: Normal
---


# SplitType property (Excel Graph)

Returns or sets the way that the two sections of either a Pie of Pie chart or a Bar of Pie chart are split. Read/write **[XlChartSplitType](excel.xlchartsplittype.md)**.

## Syntax

_expression_.**SplitType**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example must be run on either a Pie of Pie chart or a Bar of Pie chart. The example splits the two sections of the chart by value, combining all values under 10 in the primary pie and displaying them in the secondary section.

```vb
With myChart.ChartGroups(1) 
 .SplitType = xlSplitByValue 
 .SplitValue = 10 
 .VaryByCategories = True 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]