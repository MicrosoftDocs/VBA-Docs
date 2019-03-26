---
title: Axis.MinimumScaleIsAuto property (Excel)
keywords: vbaxl10.chm561091
f1_keywords:
- vbaxl10.chm561091
ms.prod: excel
api_name:
- Excel.Axis.MinimumScaleIsAuto
ms.assetid: 93767cb3-c71e-b191-2f07-7ca091498023
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.MinimumScaleIsAuto property (Excel)

 **True** if Microsoft Excel calculates the minimum value for the value axis. Read/write **Boolean**.


## Syntax

_expression_. `MinimumScaleIsAuto`

_expression_ A variable that represents an **[Axis](Excel.Axis(object).md)** object.


## Remarks

Setting the  **[MinimumScale](Excel.Axis.MinimumScale.md)** property sets this property to **False**.


## Example

This example automatically calculates the minimum scale and the maximum scale for the value axis in Chart1.


```vb
With Charts("Chart1").Axes(xlValue) 
 .MinimumScaleIsAuto = True 
 .MaximumScaleIsAuto = True 
End With
```


## See also


[Axis Object](Excel.Axis(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]