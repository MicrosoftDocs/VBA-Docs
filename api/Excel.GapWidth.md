---
title: GapWidth property (Excel Graph)
keywords: vbagr10.chm65587
f1_keywords:
- vbagr10.chm65587
ms.prod: excel
api_name:
- Excel.GapWidth
ms.assetid: d00b2a8b-76a0-1dbe-537d-bb55f3a069c9
ms.date: 04/10/2019
localization_priority: Normal
---


# GapWidth property (Excel Graph)

For Bar and Column charts: Returns or sets the space between bar or column clusters as a percentage of the bar or column width. The value of this property must be between 0 and 500. Read/write **Long**.

For Pie of Pie and Bar of Pie charts: Returns or sets the space between the primary and secondary sections of the specified chart. The value of this property must be between 5 and 200. Read/write **Long**.

## Syntax

_expression_.**GapWidth**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the space between column clusters to be 50 percent of the column width.

```vb
myChart.ChartGroups(1).GapWidth = 50
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]