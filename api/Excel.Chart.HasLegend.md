---
title: Chart.HasLegend property (Excel)
keywords: vbaxl10.chm149115
f1_keywords:
- vbaxl10.chm149115
ms.prod: excel
api_name:
- Excel.Chart.HasLegend
ms.assetid: e791cc18-03a3-1e60-f064-256cdbd6bd2e
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.HasLegend property (Excel)

 **True** if the chart has a legend. Read/write **Boolean**.


## Syntax

_expression_. `HasLegend`

_expression_ A variable that represents a [Chart](Excel.Chart-graph-object.md) object.


## Example

This example turns on the legend for Chart1 and then sets the legend font color to blue.


```vb
With Charts("Chart1") 
 .HasLegend = True 
 .Legend.Font.ColorIndex = 5 
End With
```


## See also


[Chart Object](Excel.Chart(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]