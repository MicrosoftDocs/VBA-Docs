---
title: VaryByCategories property (Excel Graph)
keywords: vbagr10.chm65596
f1_keywords:
- vbagr10.chm65596
ms.prod: excel
api_name:
- Excel.VaryByCategories
ms.assetid: e64bd5cb-1dfa-b78a-ee7e-cf3eb7b4a788
ms.date: 04/12/2019
localization_priority: Normal
---


# VaryByCategories property (Excel Graph)

**True** if Graph assigns a different color or pattern to each data marker. The chart must contain only one series. Read/write **Boolean**.


## Syntax

_expression_.**VaryByCategories**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example assigns a different color or pattern to each data marker in chart group one. The example should be run on a 2D line chart that has data markers on a series.

```vb
myChart.ChartGroups(1).VaryByCategories = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]