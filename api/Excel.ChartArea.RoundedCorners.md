---
title: ChartArea.RoundedCorners property (Excel)
keywords: vbaxl10.chm620091
f1_keywords:
- vbaxl10.chm620091
ms.prod: excel
api_name:
- Excel.ChartArea.RoundedCorners
ms.assetid: 1e9ef356-44e6-480b-bc60-a1263fd2ee90
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartArea.RoundedCorners property (Excel)

 **True** if the chart area of the chart has rounded corners. Read/write **Boolean**.


## Syntax

_expression_. `RoundedCorners`

_expression_ A variable that returns a [ChartArea](Excel.ChartArea-graph-property.md) object.


## Example

This example adds rounded corners to chart one on Sheet1.


```vb
Worksheets("Sheet1").ChartObjects(1).Chart.ChartArea.RoundedCorners = True
```


## See also


[ChartArea Object](Excel.ChartArea(object).md)

