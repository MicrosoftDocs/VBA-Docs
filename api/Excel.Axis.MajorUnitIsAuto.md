---
title: Axis.MajorUnitIsAuto property (Excel)
keywords: vbaxl10.chm561087
f1_keywords:
- vbaxl10.chm561087
ms.prod: excel
api_name:
- Excel.Axis.MajorUnitIsAuto
ms.assetid: bec8cc5a-c4c9-7d59-bf0d-ae88b9891182
ms.date: 04/13/2019
localization_priority: Normal
---


# Axis.MajorUnitIsAuto property (Excel)

**True** if Microsoft Excel calculates the major units for the value axis. Read/write **Boolean**.


## Syntax

_expression_.**MajorUnitIsAuto**

_expression_ A variable that represents an **[Axis](Excel.Axis(object).md)** object.


## Remarks

Setting the **[MajorUnit](Excel.Axis.MajorUnit.md)** property sets this property to **False**.


## Example

This example automatically sets the major and minor units for the value axis on Chart1.

```vb
With Charts("Chart1").Axes(xlValue) 
 .MajorUnitIsAuto = True 
 .MinorUnitIsAuto = True 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]