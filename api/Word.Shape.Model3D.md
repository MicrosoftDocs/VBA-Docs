---
title: Shape.Model3D property (Word)
keywords: vbaxl10.chm161480914
f1_keywords:
- vbaxl10.chm161480914
ms.prod: word
api_name:
- Word.Shape.Model3D
ms.date: 04/11/2019
localization_priority: Normal
---


# Shape.Model3D property (Word)

Returns a **[Model3DFormat](Word.Model3DFormat.md)** object that contains Model3D properties. Read-only.


## Syntax

_expression_.**Model3D**

_expression_ A variable that represents a **[Shape](Word.Shape.md)** object.


## Example

This example enables the **AutoFit** property for all Model3D objects in the active document.

```vb
For Each s In ActiveDocument.Shapes
 If s.Type = mso3DModel Then s.Model3D.AutoFit = True
Next
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]