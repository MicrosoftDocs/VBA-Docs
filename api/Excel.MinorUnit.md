---
title: MinorUnit property (Excel Graph)
keywords: vbagr10.chm3077551
f1_keywords:
- vbagr10.chm3077551
ms.prod: excel
api_name:
- Excel.MinorUnit
ms.assetid: 9da86e1c-dfc2-49c8-e6bd-1e5529b2da33
ms.date: 04/11/2019
localization_priority: Normal
---


# MinorUnit property (Excel Graph)

Returns or sets the minor units on the axis. Read/write **Double**.

## Syntax

_expression_.**MinorUnit**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Remarks

Setting this property sets the **[MinorUnitIsAuto](Excel.MinorUnitIsAuto.md)** property to **False**.

Use the **[TickMarkSpacing](Excel.TickMarkSpacing.md)** property to set tick-mark spacing on the category axis.


## Example

This example sets the major and minor units for the value axis.

```vb
With myChart.Axes(xlValue) 
 .MajorUnit = 100 
 .MinorUnit = 20 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]