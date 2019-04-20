---
title: ChartGroup.SplitType property (Excel)
keywords: vbaxl10.chm568097
f1_keywords:
- vbaxl10.chm568097
ms.prod: excel
api_name:
- Excel.ChartGroup.SplitType
ms.assetid: c65ca7a4-59b1-6b15-116a-f76007fbd4be
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartGroup.SplitType property (Excel)

Returns or sets the way the two sections of either a Pie of Pie chart or a Bar of Pie chart are split. Read/write **[XlChartSplitType](Excel.XlChartSplitType.md)**.


## Syntax

_expression_.**SplitType**

_expression_ A variable that represents a **[ChartGroup](Excel.ChartGroup(object).md)** object.


## Example

This example must be run on either a Pie of Pie chart or a Bar of Pie chart. The example splits the two sections of the chart by value, combining all values under 10 in the primary pie and displaying them in the secondary section.

```vb
With Worksheets(1).ChartObjects(1).Chart.ChartGroups(1) 
 .SplitType = xlSplitByValue 
 .SplitValue = 10 
 .VaryByCategories = True 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]