---
title: TextColumn.SpaceAfter property (Word)
ms.prod: word
api_name:
- Word.TextColumn.SpaceAfter
ms.assetid: 95b77d91-e13a-c6d3-f8c3-069c81b39cb1
ms.date: 06/08/2017
localization_priority: Normal
---


# TextColumn.SpaceAfter property (Word)

Returns or sets the amount of spacing (in points) after the specified paragraph or text column. Read/write  **Single**.


## Syntax

_expression_. `SpaceAfter`

_expression_ Required. A variable that represents a '[TextColumn](Word.TextColumn.md)' object.


## Example

This example sets the active document to three columns with a 0.5-inch space after the first column. The  **InchesToPoints** method is used to convert inches to points.


```vb
With ActiveDocument.PageSetup.TextColumns 
 .SetCount NumColumns:=3 
 .LineBetween = False 
 .EvenlySpaced = True 
 .Item(1).SpaceAfter = InchesToPoints(0.5) 
End With
```


## See also


[TextColumn Object](Word.TextColumn.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]