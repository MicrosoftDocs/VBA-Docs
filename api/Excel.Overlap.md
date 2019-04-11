---
title: Overlap property (Excel Graph)
keywords: vbagr10.chm5207749
f1_keywords:
- vbagr10.chm5207749
ms.prod: excel
api_name:
- Excel.Overlap
ms.assetid: 60e82754-4553-7ee9-7403-06cd12de733e
ms.date: 04/11/2019
localization_priority: Normal
---


# Overlap property (Excel Graph)

Specifies how bars and columns are positioned. Can be a value between -100 and 100. Applies only to 2D bar and 2D column charts. Read/write **Long**.

## Syntax

_expression_.**Overlap**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Remarks

If this property is set to -100, bars are positioned so that there's one bar width between them. If the overlap is 0 (zero), there's no space between bars (one bar starts immediately after the preceding bar). If the overlap is 100, bars are positioned on top of each other.


## Example

This example sets the overlap for chart group one to -50. The example should be run on a 2D column chart that has two or more series.

```vb
myChart.ChartGroups(1).Overlap = -50
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]