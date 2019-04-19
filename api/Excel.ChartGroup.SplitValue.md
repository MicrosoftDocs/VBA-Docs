---
title: ChartGroup.SplitValue property (Excel)
keywords: vbaxl10.chm568098
f1_keywords:
- vbaxl10.chm568098
ms.prod: excel
api_name:
- Excel.ChartGroup.SplitValue
ms.assetid: a7cab670-1510-5334-f11b-12dc8cc13570
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartGroup.SplitValue property (Excel)

Returns or sets the threshold value separating the two sections of either a Pie of Pie chart or a Bar of Pie chart. Read/write **Variant**.


## Syntax

_expression_.**SplitValue**

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