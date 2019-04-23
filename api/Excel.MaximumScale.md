---
title: MaximumScale property (Excel Graph)
keywords: vbagr10.chm5207676
f1_keywords:
- vbagr10.chm5207676
ms.prod: excel
api_name:
- Excel.MaximumScale
ms.assetid: 1fd6633e-7782-78d0-ba24-9c3d46f85471
ms.date: 04/11/2019
localization_priority: Normal
---


# MaximumScale property (Excel Graph)

Returns or sets the maximum value on the axis. Read/write **Double**.

## Syntax

_expression_.**MaximumScale**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Remarks

Setting this property sets the **[MaximumScaleIsAuto](Excel.MaximumScaleIsAuto.md)** property to **False**.


## Example

This example sets the minimum and maximum values for the value axis.

```vb
With myChart.Axes(xlValue) 
 .MinimumScale = 10 
 .MaximumScale = 120 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]