---
title: MinorUnitIsAuto property (Excel Graph)
keywords: vbagr10.chm65576
f1_keywords:
- vbagr10.chm65576
ms.prod: excel
api_name:
- Excel.MinorUnitIsAuto
ms.assetid: ca6a18d5-f93f-4801-7704-4d3a25b633cb
ms.date: 04/11/2019
localization_priority: Normal
---


# MinorUnitIsAuto property (Excel Graph)

**True** if Graph calculates minor units for the axis. Read/write **Boolean**.


## Syntax

_expression_.**MinorUnitIsAuto**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

Setting the **[MinorUnit](Excel.MinorUnit.md)** property sets this property to **False**.


## Example

This example automatically calculates major and minor units for the value axis.

```vb
With myChart.Axes(xlValue) 
 .MajorUnitIsAuto = True 
 .MinorUnitIsAuto = True 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]