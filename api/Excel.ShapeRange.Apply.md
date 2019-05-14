---
title: ShapeRange.Apply method (Excel)
keywords: vbaxl10.chm640078
f1_keywords:
- vbaxl10.chm640078
ms.prod: excel
api_name:
- Excel.ShapeRange.Apply
ms.assetid: 34acef44-7075-ffc1-199c-3396e17caafe
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeRange.Apply method (Excel)

Applies to the specified shape formatting that's been copied by using the  **[PickUp](Excel.ShapeRange.PickUp.md)** method.


## Syntax

_expression_.**Apply**

_expression_ A variable that represents a **[ShapeRange](Excel.shaperange.md)** object.


## Example

This example copies the formatting of shape one on _myDocument_ and then applies the copied formatting to shape two.


```vb
Set myDocument = Worksheets(1) 
With myDocument 
 .Shapes(1).PickUp 
 .Shapes(2).Apply 
End With
```


## See also


[ShapeRange Object](Excel.ShapeRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]