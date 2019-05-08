---
title: Axis.MaximumScaleIsAuto property (Excel)
keywords: vbaxl10.chm561089
f1_keywords:
- vbaxl10.chm561089
ms.prod: excel
api_name:
- Excel.Axis.MaximumScaleIsAuto
ms.assetid: c0e0f4b6-5d1c-5acb-2e7a-8722e10cd2bc
ms.date: 04/13/2019
localization_priority: Normal
---


# Axis.MaximumScaleIsAuto property (Excel)

**True** if Microsoft Excel calculates the maximum value for the value axis. Read/write **Boolean**.


## Syntax

_expression_.**MaximumScaleIsAuto**

_expression_ A variable that represents an **[Axis](Excel.Axis(object).md)** object.


## Remarks

Setting the **[MaximumScale](Excel.Axis.MaximumScale.md)** property sets this property to **False**.


## Example

This example automatically calculates the minimum scale and the maximum scale for the value axis on Chart1.

```vb
With Charts("Chart1").Axes(xlValue) 
 .MinimumScaleIsAuto = True 
 .MaximumScaleIsAuto = True 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]