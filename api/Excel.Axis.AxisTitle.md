---
title: Axis.AxisTitle property (Excel)
keywords: vbaxl10.chm561075
f1_keywords:
- vbaxl10.chm561075
ms.prod: excel
api_name:
- Excel.Axis.AxisTitle
ms.assetid: 33ba6b94-189b-e9d0-a153-af028380a58a
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.AxisTitle property (Excel)

Returns an  **[AxisTitle](Excel.AxisTitle(object).md)** object that represents the title of the specified axis. Read-only.


## Syntax

_expression_. `AxisTitle`

_expression_ A variable that represents an [Axis](Excel.Axis-graph-object.md) object.


## Remarks

This example adds an axis label to the category axis in Chart1.


## Example


```vb
With Charts("Chart1").Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Text = "July Sales" 
End With
```


## See also


[Axis Object](Excel.Axis(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]