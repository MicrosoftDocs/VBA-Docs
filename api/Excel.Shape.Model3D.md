---
title: Shape.Model3D property (Excel)
keywords: vbaxl10.chm636158
f1_keywords:
- vbaxl10.chm636158
ms.prod: excel
api_name:
- Excel.Shape.Model3D
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.Model3D property (Excel)

Returns a **[Model3DFormat](Excel.Model3DFormat.md)** object that contains Model3D properties. Read-only.


## Syntax

_expression_.**Model3D**

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


## Example

This example enables the **[AutoFit](Excel.Model3DFormat.AutoFit.md)** property for all Model3D objects on worksheet one.

```vb
For Each s In Worksheets(1).Shapes
 If s.Type = mso3DModel Then s.Model3D.AutoFit = True
Next
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]