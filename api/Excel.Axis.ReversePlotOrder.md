---
title: Axis.ReversePlotOrder property (Excel)
keywords: vbaxl10.chm561096
f1_keywords:
- vbaxl10.chm561096
ms.prod: excel
api_name:
- Excel.Axis.ReversePlotOrder
ms.assetid: 151f544c-4e6d-a583-5fab-5df0e5269681
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.ReversePlotOrder property (Excel)

 **True** if Microsoft Excel plots data points from last to first. Read/write **Boolean**.


## Syntax

_expression_. `ReversePlotOrder`

_expression_ A variable that represents an [Axis](Excel.Axis-graph-object.md) object.


## Remarks

This property cannot be used on radar charts.


## Example

This example plots data points from last to first on the value axis on Chart1.


```vb
Charts("Chart1").Axes(xlValue).ReversePlotOrder = True
```


## See also


[Axis Object](Excel.Axis(object).md)

