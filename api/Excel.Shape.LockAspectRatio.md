---
title: Shape.LockAspectRatio property (Excel)
keywords: vbaxl10.chm636102
f1_keywords:
- vbaxl10.chm636102
ms.prod: excel
api_name:
- Excel.Shape.LockAspectRatio
ms.assetid: 1b517827-ebe0-a6ae-0fd7-fe3049eb6d04
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.LockAspectRatio property (Excel)

**True** if the specified shape retains its original proportions when you resize it. **False** if you can change the height and width of the shape independently of one another when you resize it. Read/write **[MsoTriState](Office.MsoTriState.md)**.


## Syntax

_expression_.**LockAspectRatio**

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.



## Example

This example adds a cube to _myDocument_. The cube can be moved and resized, but not reproportioned.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.AddShape(msoShapeCube, _ 
    50, 50, 100, 200).LockAspectRatio = msoTrue
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]