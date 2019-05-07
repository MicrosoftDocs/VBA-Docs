---
title: Document.GridDistanceVertical property (Word)
keywords: vbawd10.chm158007599
f1_keywords:
- vbawd10.chm158007599
ms.prod: word
api_name:
- Word.Document.GridDistanceVertical
ms.assetid: 4b3c6f15-a379-9399-fab6-ac6ec45717fa
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.GridDistanceVertical property (Word)

Returns or sets a  **Single** that represents the amount of vertical space between the invisible gridlines that Microsoft Word uses when you draw, move, and resize AutoShapes or East Asian characters in the specified document. Read/write.


## Syntax

_expression_. `GridDistanceVertical`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Example

This example sets the horizontal and vertical distance between gridlines and then enables the Snap objects to grid feature for the current document.


```vb
With ActiveDocument 
 .GridDistanceHorizontal = 9 
 .GridDistanceVertical = 9 
 .SnapToGrid = True 
End With
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]