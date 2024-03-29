---
title: Crosses property (Excel Graph)
keywords: vbagr10.chm3076987
f1_keywords:
- vbagr10.chm3076987
api_name:
- Excel.Crosses
ms.assetid: 60c2ae55-87ad-f28d-5739-cbd51c8144be
ms.date: 04/10/2019
ms.localizationpriority: medium
---


# Crosses property (Excel Graph)

Returns or sets the point on the specified axis where the other axis crosses. Read/write **[XlAxisCrosses](excel.xlaxiscrosses.md)**.

## Syntax

_expression_.**Crosses**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

This property isn't available for radar charts. For 3D charts, this property indicates where the plane defined by the category axis crosses the value axis.

This property can be used for both category and value axes. On the category axis, **xlMinimum** sets the value axis to cross at the first category, and **xlMaximum** sets the value axis to cross at the last category.

Note that **xlMinimum** and **xlMaximum** can have different meanings, depending on the axis.


## Example

This example sets the value axis to cross the category axis at the maximum x value.

```vb
myChart.Axes(xlCategory).Crosses = xlMaximum
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]