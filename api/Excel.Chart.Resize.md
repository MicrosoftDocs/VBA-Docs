---
title: Chart.Resize event (Excel)
keywords: vbaxl10.chm500075
f1_keywords:
- vbaxl10.chm500075
ms.prod: excel
api_name:
- Excel.Chart.Resize
ms.assetid: d1b7d0bb-d190-18f2-83f9-b91b637d80aa
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.Resize event (Excel)

Occurs when the chart is resized.


## Syntax

_expression_.**Resize**

_expression_ A variable that returns a **[Chart](Excel.Chart(object).md)** object.


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




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]