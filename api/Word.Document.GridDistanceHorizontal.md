---
title: Document.GridDistanceHorizontal property (Word)
keywords: vbawd10.chm158007598
f1_keywords:
- vbawd10.chm158007598
ms.prod: word
api_name:
- Word.Document.GridDistanceHorizontal
ms.assetid: dabff5b7-420c-ffb7-1812-eeadbdacc864
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.GridDistanceHorizontal property (Word)

Returns or sets a  **Single** that represents the amount of horizontal space between the invisible gridlines that Microsoft Word uses when you draw, move, and resize AutoShapes or East Asian characters in the specified document. Read/write.


## Syntax

_expression_. `GridDistanceHorizontal`

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