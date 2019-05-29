---
title: Options.GridOriginHorizontal property (Word)
keywords: vbawd10.chm162988115
f1_keywords:
- vbawd10.chm162988115
ms.prod: word
api_name:
- Word.Options.GridOriginHorizontal
ms.assetid: b364fde9-c889-e139-49eb-91fdff42ac96
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.GridOriginHorizontal property (Word)

Returns or sets the point, relative to the left edge of the page, where you want the invisible grid for drawing, moving, and resizing AutoShapes or East Asian characters to begin in new documents. Read/write  **Single**.


## Syntax

_expression_. `GridOriginHorizontal`

_expression_ A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example sets the horizontal and vertical point of origin for the grid, sets the horizontal and vertical distance between gridlines, and then enables the Snap objects to grid feature for a new document.


```vb
With Options 
 .GridOriginHorizontal = InchesToPoints(1) 
 .GridOriginVertical = InchesToPoints(2) 
 .GridDistanceHorizontal = InchesToPoints(0.1) 
 .GridDistanceVertical = InchesToPoints(0.1) 
 .SnapToGrid = True 
End With 
Documents.Add
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]