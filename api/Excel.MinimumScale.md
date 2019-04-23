---
title: MinimumScale property (Excel Graph)
keywords: vbagr10.chm65569
f1_keywords:
- vbagr10.chm65569
ms.prod: excel
api_name:
- Excel.MinimumScale
ms.assetid: 4aca27ef-c1af-e74e-8ca5-6a3fc1aefaa2
ms.date: 04/11/2019
localization_priority: Normal
---


# MinimumScale property (Excel Graph)

Returns or sets the minimum value on the axis. Read/write **Double**.

## Syntax

_expression_.**MinimumScale**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

Setting this property sets the **[MinimumScaleIsAuto](Excel.MinimumScaleIsAuto.md)** property to **False**.


## Example

This example sets the minimum and maximum values for the value axis.

```vb
With myChart.Axes(xlValue) 
 .MinimumScale = 10 
 .MaximumScale = 120 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]