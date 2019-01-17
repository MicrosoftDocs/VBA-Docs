---
title: Chart.Resize Event (Excel)
keywords: vbaxl10.chm500075
f1_keywords:
- vbaxl10.chm500075
ms.prod: excel
api_name:
- Excel.Chart.Resize
ms.assetid: d1b7d0bb-d190-18f2-83f9-b91b637d80aa
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.Resize Event (Excel)

Occurs when the chart is resized.


## Syntax

_expression_. `Resize`

_expression_ A variable that returns a '[Chart](Excel.Chart(object).md)' object.


## Example

The following code example keeps the upper-left corner of the chart at the same location when the chart is resized.


```vb
Private Sub myChartClass_Resize() 
 With ActiveChart.Parent 
 .Left = 100 
 .Top = 150 
 End With 
End Sub
```


## See also


[Chart Object](Excel.Chart(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]