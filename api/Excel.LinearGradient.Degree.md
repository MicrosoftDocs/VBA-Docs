---
title: LinearGradient.Degree property (Excel)
keywords: vbaxl10.chm855074
f1_keywords:
- vbaxl10.chm855074
ms.prod: excel
api_name:
- Excel.LinearGradient.Degree
ms.assetid: 0608fe59-76e9-e199-2cc6-848f283813f3
ms.date: 04/30/2019
localization_priority: Normal
---


# LinearGradient.Degree property (Excel)

The angle of the linear gradient fill within a selection. Read/write.


## Syntax

_expression_.**Degree**

_expression_ A variable that represents a **[LinearGradient](Excel.LinearGradient.md)** object.


## Return value

Double


## Remarks

Uses values ranging from 0&ndash;360.


## Example

```vb
With Selection.Interior 
 .Pattern = xlPatternLinearGradient 
 .Gradient.Degree = 45 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]