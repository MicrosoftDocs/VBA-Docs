---
title: Shape.LockAspectRatio property (Word)
keywords: vbawd10.chm161480817
f1_keywords:
- vbawd10.chm161480817
ms.prod: word
api_name:
- Word.Shape.LockAspectRatio
ms.assetid: dd408737-405f-4b91-0eae-73161fe38425
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.LockAspectRatio property (Word)

 **MsoTrue** if the specified shape retains its original proportions when you resize it. **MsoFalse** if you can change the height and width of the shape independently of one another when you resize it. Read/write **MsoTriState**.


## Syntax

_expression_.**LockAspectRatio**

_expression_ Required. A variable that represents a **[Shape](Word.Shape.md)** object.


## Example

This example adds a cube to _myDocument_. The cube can be moved and resized but not reproportioned.


```vb
Set myDocument = ActiveDocument 
myDocument.Shapes.AddShape(msoShapeCube, _ 
 50, 50, 100, 200).LockAspectRatio = msoTrue
```


## See also


[Shape Object](Word.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]