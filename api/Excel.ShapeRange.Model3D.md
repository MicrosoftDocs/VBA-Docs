---
title: ShapeRange.Model3D property (Excel)
keywords: vbaxl10.chm640148
f1_keywords:
- vbaxl10.chm640148
ms.prod: excel
api_name:
- Excel.ShapeRange.Model3D
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeRange.Model3D property (Excel)

Returns a **[Model3DFormat](Excel.Model3DFormat.md)** object that contains Model3D properties. Read-only.


## Syntax

_expression_.**Model3D**

_expression_ A variable that represents a **[ShapeRange](Excel.ShapeRange.md)** object.


## Example

This example selects all shapes in worksheet one, and then disables the **[AutoFit](Excel.Model3DFormat.AutoFit.md)** property of all Model3D objects in the selection.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.SelectAll 
Selection.ShapeRange.Model3D.AutoFit = False
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]