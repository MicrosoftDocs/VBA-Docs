---
title: Shape.Apply method (Excel)
keywords: vbaxl10.chm636074
f1_keywords:
- vbaxl10.chm636074
ms.prod: excel
api_name:
- Excel.Shape.Apply
ms.assetid: fe094baf-76d7-8418-aa34-c90d37f95def
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.Apply method (Excel)

Applies to the specified shape formatting that's been copied by using the **[PickUp](Excel.Shape.PickUp.md)** method.


## Syntax

_expression_.**Apply**

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


## Example

This example copies the formatting of shape one on _myDocument_ and then applies the copied formatting to shape two.

```vb
Set myDocument = Worksheets(1) 
With myDocument 
 .Shapes(1).PickUp 
 .Shapes(2).Apply 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]