---
title: Axis.Crosses property (Excel)
keywords: vbaxl10.chm561078
f1_keywords:
- vbaxl10.chm561078
api_name:
- Excel.Axis.Crosses
ms.assetid: 571e256d-b711-e3cd-f0f2-c53e86375e6f
ms.date: 04/13/2019
ms.localizationpriority: medium
---


# Axis.Crosses property (Excel)

Returns or sets the point on the specified axis where the other axis crosses. Read/write **Long**.


## Syntax

_expression_.**Crosses**

_expression_ A variable that represents an **[Axis](Excel.Axis(object).md)** object.


## Remarks

Can be one of the **[XlAxisCrosses](excel.xlaxiscrosses.md)** constants.

This property isn't available for radar charts. For 3D charts, this property can only be applied to the value axis, and indicates where the plane defined by the category axes crosses the value axis.

This property can be used for both category and value axes. On the category axis, **xlMinimum** sets the value axis to cross at the first category, and **xlMaximum** sets the value axis to cross at the last category.

Note that **xlMinimum** and **xlMaximum** can have different meanings, depending on the axis.


## Example

This example sets the value axis on Chart1 to cross the category axis at the maximum x value.

```vb
Charts("Chart1").Axes(xlCategory).Crosses = xlMaximum
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]