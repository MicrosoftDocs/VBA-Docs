---
title: Axis.AxisBetweenCategories property (Excel)
keywords: vbaxl10.chm561073
f1_keywords:
- vbaxl10.chm561073
ms.prod: excel
api_name:
- Excel.Axis.AxisBetweenCategories
ms.assetid: 03a2d87b-1fbd-470d-01d2-e4156dae55e2
ms.date: 04/13/2019
localization_priority: Normal
---


# Axis.AxisBetweenCategories property (Excel)

**True** if the value axis crosses the category axis between categories. Read/write **Boolean**.


## Syntax

_expression_.**AxisBetweenCategories**

_expression_ A variable that represents an **[Axis](Excel.Axis(object).md)** object.


## Remarks

This property applies only to category axes, and it doesn't apply to 3D charts.


## Example

This example causes the value axis on Chart1 to cross the category axis between categories.

```vb
Charts("Chart1").Axes(xlCategory).AxisBetweenCategories = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]