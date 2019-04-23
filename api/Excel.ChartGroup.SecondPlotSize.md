---
title: ChartGroup.SecondPlotSize property (Excel)
keywords: vbaxl10.chm568099
f1_keywords:
- vbaxl10.chm568099
ms.prod: excel
api_name:
- Excel.ChartGroup.SecondPlotSize
ms.assetid: 231541fa-0353-3533-6b4b-0653b6157568
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartGroup.SecondPlotSize property (Excel)

Returns or sets the size of the secondary section of either a Pie of Pie chart or a Bar of Pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200. Read/write **Long**.


## Syntax

_expression_.**SecondPlotSize**

_expression_ A variable that represents a **[ChartGroup](Excel.ChartGroup(object).md)** object.


## Example

This example must be run on either a Pie of Pie chart or a Bar of Pie chart. The example splits the two sections of the chart by value, combining all values under 10 in the primary pie and displaying them in the secondary section. The secondary section is 50 percent of the size of the primary pie.

```vb
With Worksheets(1).ChartObjects(1).Chart.ChartGroups(1) 
 .SplitType = xlSplitByValue 
 .SplitValue = 10 
 .VaryByCategories = True 
 .SecondPlotSize = 50 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]