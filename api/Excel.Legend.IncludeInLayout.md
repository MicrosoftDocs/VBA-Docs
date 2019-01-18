---
title: Legend.IncludeInLayout property (Excel)
keywords: vbaxl10.chm622090
f1_keywords:
- vbaxl10.chm622090
ms.prod: excel
api_name:
- Excel.Legend.IncludeInLayout
ms.assetid: ebb55dfa-8b3e-b247-4574-65b22640eadd
ms.date: 06/08/2017
localization_priority: Normal
---


# Legend.IncludeInLayout property (Excel)

 **True** if a legend will occupy the chart layout space when a chart layout is being determined. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_. `IncludeInLayout`

_expression_ A variable that represents a [Legend](Excel.Legend-graph-property.md) object.


## Remarks

This property does not affect whether a chart is in autolayout mode or not. If the user adds a title using the  **Above Chart** command, the chart will resize smaller. If the user then removes the title or selects one of the overlay title options, the chart will resize larger, as if the title were not on the chart.


## See also


[Legend Object](Excel.Legend(object).md)

