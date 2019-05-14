---
title: Shape.PickUp method (Excel)
keywords: vbaxl10.chm636081
f1_keywords:
- vbaxl10.chm636081
ms.prod: excel
api_name:
- Excel.Shape.PickUp
ms.assetid: 77da5d6d-35f8-71c3-70ee-481f59c5674b
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.PickUp method (Excel)

Copies the formatting of the specified shape. Use the **[Apply](Excel.Shape.Apply.md)** method to apply the copied formatting to another shape.


## Syntax

_expression_.**PickUp**

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