---
title: Cells.Width property (Word)
keywords: vbawd10.chm155844614
f1_keywords:
- vbawd10.chm155844614
ms.prod: word
api_name:
- Word.Cells.Width
ms.assetid: e46b835d-3fbd-8149-9fbb-00c40ffc0ff5
ms.date: 06/08/2017
localization_priority: Normal
---


# Cells.Width property (Word)

Returns or sets the width of the table cells, in points. Read/write  **Long**.


## Syntax

_expression_.**Width**

_expression_ A variable that represents a '[Cells](Word.cells.md)' object.


## Example

This example returns the width (in inches) of the cells within the selection.


```vb
If Selection.Information(wdWithInTable) = True Then 
 MsgBox PointsToInches(Selection.Cells.Width) 
End If
```


## See also


[Cells Collection Object](Word.cells.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]