---
title: ChartGroup.SizeRepresents property (Excel)
keywords: vbaxl10.chm568094
f1_keywords:
- vbaxl10.chm568094
ms.prod: excel
api_name:
- Excel.ChartGroup.SizeRepresents
ms.assetid: db7811b5-6d65-d3e0-0c0b-83dcd3692cf1
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroup.SizeRepresents property (Excel)

Returns or sets what the bubble size represents on a bubble chart. Can be either of the following  **[xlSizeRepresents](Excel.XlSizeRepresents.md)** constants: **xlSizeIsArea** or **xlSizeIsWidth**. Read/write **Long**.


## Syntax

_expression_. `SizeRepresents`

_expression_ A variable that represents a [ChartGroup](Excel.ChartGroup-graph-object.md) object.


## Example

This example sets what the bubble size represents for chart group one.


```vb
Charts(1).ChartGroups(1).SizeRepresents = xlSizeIsWidth
```


## See also


[ChartGroup Object](Excel.ChartGroup(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]