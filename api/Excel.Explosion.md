---
title: Explosion property (Excel Graph)
ms.prod: excel
api_name:
- Excel.Explosion
ms.assetid: 252a3533-28df-4317-8af1-7509339409a5
ms.date: 04/10/2019
localization_priority: Normal
---


# Explosion property (Excel Graph)

Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie). Read/write **Long**.

## Syntax

_expression_.**Explosion**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the explosion value for point two. The example should be run on a pie chart.

```vb
myChart.SeriesCollection(1).Points(2). Explosion = 20

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]