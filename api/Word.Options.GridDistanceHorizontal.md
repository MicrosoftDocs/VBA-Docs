---
title: Options.GridDistanceHorizontal property (Word)
keywords: vbawd10.chm162988113
f1_keywords:
- vbawd10.chm162988113
ms.prod: word
api_name:
- Word.Options.GridDistanceHorizontal
ms.assetid: 1d28ba4b-ee06-1b1a-e921-2d8d07cab305
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.GridDistanceHorizontal property (Word)

Returns or sets the amount of horizontal space between the invisible gridlines that Word uses when you draw, move, and resize AutoShapes or East Asian characters in new documents. Read/write  **Single**.


## Syntax

_expression_. `GridDistanceHorizontal`

_expression_ A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example sets the horizontal and vertical distance between gridlines and then enables the Snap objects to grid feature for a new document.


```vb
With Options 
 .GridDistanceHorizontal = InchesToPoints(0.2) 
 .GridDistanceVertical = InchesToPoints(0.2) 
 .SnapToGrid = True 
End With 
Documents.Add
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]