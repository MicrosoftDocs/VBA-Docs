---
title: Chart.PlotVisibleOnly property (Excel)
keywords: vbaxl10.chm149135
f1_keywords:
- vbaxl10.chm149135
ms.prod: excel
api_name:
- Excel.Chart.PlotVisibleOnly
ms.assetid: e09aee43-c3f7-9269-f01a-d6298ab780fa
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.PlotVisibleOnly property (Excel)

 **True** if only visible cells are plotted. **False** if both visible and hidden cells are plotted. Read/write **Boolean**.


## Syntax

_expression_. `PlotVisibleOnly`

_expression_ A variable that represents a [Chart](Excel.Chart-graph-object.md) object.


## Example

This example causes Microsoft Excel to plot only visible cells in Chart1.


```vb
Charts("Chart1").PlotVisibleOnly = True
```


## See also


[Chart Object](Excel.Chart(object).md)

