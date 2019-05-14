---
title: ShapeRange.LockAspectRatio property (Excel)
keywords: vbaxl10.chm640109
f1_keywords:
- vbaxl10.chm640109
ms.prod: excel
api_name:
- Excel.ShapeRange.LockAspectRatio
ms.assetid: 58b33bc9-de5c-1fb2-7369-7f4f8dedde58
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeRange.LockAspectRatio property (Excel)

**True** if the specified shape retains its original proportions when you resize it. **False** if you can change the height and width of the shape independently of one another when you resize it. Read/write **[MsoTriState](Office.MsoTriState.md)**.


## Syntax

_expression_.**LockAspectRatio**

_expression_ A variable that represents a **[ShapeRange](Excel.shaperange.md)** object.


## Example

This example adds a cube to _myDocument_. The cube can be moved and resized, but not reproportioned.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.AddShape(msoShapeCube, _ 
    50, 50, 100, 200).LockAspectRatio = msoTrue
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]