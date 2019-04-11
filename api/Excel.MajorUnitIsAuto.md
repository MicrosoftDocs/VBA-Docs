---
title: MajorUnitIsAuto property (Excel Graph)
keywords: vbagr10.chm5207645
f1_keywords:
- vbagr10.chm5207645
ms.prod: excel
api_name:
- Excel.MajorUnitIsAuto
ms.assetid: 6eda8012-2ef3-d23b-bace-e2695a5e80f5
ms.date: 04/11/2019
localization_priority: Normal
---


# MajorUnitIsAuto property (Excel Graph)

**True** if Graph calculates the major units for the axis. Read/write **Boolean**.

## Syntax

_expression_.**MajorUnitIsAuto**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

Setting the **[MajorUnit](Excel.MajorUnit.md)** property sets this property to **False**.

## Example

This example automatically sets the major and minor units for the value axis.

```vb
With myChart.Axes(xlValue) 
 .MajorUnitIsAuto = True 
 .MinorUnitIsAuto = True 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]