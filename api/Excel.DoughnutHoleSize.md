---
title: DoughnutHoleSize property (Excel Graph)
keywords: vbagr10.chm66662
f1_keywords:
- vbagr10.chm66662
ms.prod: excel
api_name:
- Excel.DoughnutHoleSize
ms.assetid: 07e1e63b-8e31-92e5-18ab-c47104d093ac
ms.date: 04/10/2019
localization_priority: Normal
---


# DoughnutHoleSize property (Excel Graph)

Returns or sets the size of the hole in a doughnut chart group. The hole size is expressed as a percentage of the chart size, between 10 and 90 percent. Read/write **Long**.

## Syntax

_expression_.**DoughnutHoleSize**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the hole size for doughnut group one. The example should be run on a 2D doughnut chart.

```vb
myChart.DoughnutGroups(1).DoughnutHoleSize = 10
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]