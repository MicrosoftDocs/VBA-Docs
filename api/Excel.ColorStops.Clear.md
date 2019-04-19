---
title: ColorStops.Clear method (Excel)
keywords: vbaxl10.chm853078
f1_keywords:
- vbaxl10.chm853078
ms.prod: excel
api_name:
- Excel.ColorStops.Clear
ms.assetid: 308edcb7-6085-77d6-5e6a-d8ec1d31c043
ms.date: 04/20/2019
localization_priority: Normal
---


# ColorStops.Clear method (Excel)

Clears the represented object.


## Syntax

_expression_.**Clear**

_expression_ An expression that returns a **[ColorStops](Excel.ColorStops.md)** object.


## Return value

Nothing


## Example

Clears the current **ColorStops**.

```vb
Range("A1:A10").Select 
With Selection.Interior 
 .Pattern = xlPatternLinearGradient 
 .Gradient.Degree = 90 
 .Gradient.ColorStops.Clear 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]