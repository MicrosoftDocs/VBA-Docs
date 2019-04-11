---
title: SplitValue property (Excel Graph)
keywords: vbagr10.chm5208024
f1_keywords:
- vbagr10.chm5208024
ms.prod: excel
api_name:
- Excel.SplitValue
ms.assetid: 3200801a-9464-6bde-59a2-0a8baafcb8ff
ms.date: 04/12/2019
localization_priority: Normal
---


# SplitValue property (Excel Graph)

Returns or sets the threshold value separating the two sections of either a Pie of Pie chart or a Bar of Pie chart. Read/write **Variant**.


## Syntax

_expression_.**SplitValue**

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