---
title: BarShape property (Excel Graph)
keywords: vbagr10.chm66939
f1_keywords:
- vbagr10.chm66939
ms.prod: excel
api_name:
- Excel.BarShape
ms.assetid: 2da9b9aa-84db-6ade-845e-abcb142acc3b
ms.date: 04/09/2019
localization_priority: Normal
---


# BarShape property (Excel Graph)

Returns or sets the shape used with the specified 3D bar or column chart. Read/write **[XlBarShape](excel.xlbarshape.md)**.

## Syntax

_expression_.**BarShape**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the shape used with series one on the chart.

```vb
myChart.SeriesCollection(1).BarShape = xlConeToPoint
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]