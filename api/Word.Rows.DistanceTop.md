---
title: Rows.DistanceTop property (Word)
keywords: vbawd10.chm155975693
f1_keywords:
- vbawd10.chm155975693
ms.prod: word
api_name:
- Word.Rows.DistanceTop
ms.assetid: 50ff15c4-708b-d8a1-9040-83f59dcf766c
ms.date: 06/08/2017
localization_priority: Normal
---


# Rows.DistanceTop property (Word)

Returns or sets the distance (in points) between the document text and the top edge of the specified table. Read/write  **Single**.


## Syntax

_expression_. `DistanceTop`

_expression_ A variable that represents a **[Rows](Word.Rows.md)** object.


## Remarks

This property doesn't have any effect if  **WrapAroundText** is **False**.


## Example

This example sets text to wrap around the first table in the active document and sets the distance for wrapped text to 20 points on all sides of the table.


```vb
With ActiveDocument.Tables(1).Rows 
 .WrapAroundText = True 
 .DistanceLeft = 20 
 .DistanceRight = 20 
 .DistanceTop = 20 
 .DistanceBottom = 20 
End With
```


## See also


[Rows Collection Object](Word.rows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]