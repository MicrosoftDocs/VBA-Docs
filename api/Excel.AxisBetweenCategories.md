---
title: AxisBetweenCategories property (Excel Graph)
keywords: vbagr10.chm65581
f1_keywords:
- vbagr10.chm65581
api_name:
- Excel.AxisBetweenCategories
ms.assetid: 4ca52b75-036d-0851-c3cd-aa2deca0907e
ms.date: 04/09/2019
ms.localizationpriority: medium
---


# AxisBetweenCategories property (Excel Graph)

**True** if the value axis crosses the category axis between categories. Read/write **Boolean**.

## Syntax

_expression_.**AxisBetweenCategories**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Remarks

This property applies only to category axes, and it doesn't apply to 3D charts.


## Example

This example causes the value axis to cross the category axis between categories.

```vb
myChart.Axes(xlCategory).AxisBetweenCategories = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]