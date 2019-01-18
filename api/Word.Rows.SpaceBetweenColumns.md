---
title: Rows.SpaceBetweenColumns property (Word)
keywords: vbawd10.chm155975686
f1_keywords:
- vbawd10.chm155975686
ms.prod: word
api_name:
- Word.Rows.SpaceBetweenColumns
ms.assetid: 286e0236-eab3-18d2-926a-d27e2516e62b
ms.date: 06/08/2017
localization_priority: Normal
---


# Rows.SpaceBetweenColumns property (Word)

Returns or sets the distance (in points) between text in adjacent columns of the specified row or rows. Read/write  **Single**.


## Syntax

 _expression_. `SpaceBetweenColumns`

 _expression_ Required. A variable that represents a '[Rows](Word.rows.md)' collection.


## Example

This example returns the distance (in points) between columns in the selected table rows.


```vb
If Selection.Information(wdWithInTable) = True Then 
 MsgBox Selection.Rows.SpaceBetweenColumns 
End If
```


## See also


[Rows Collection Object](Word.rows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]