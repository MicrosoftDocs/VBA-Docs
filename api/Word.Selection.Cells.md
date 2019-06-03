---
title: Selection.Cells property (Word)
keywords: vbawd10.chm158662713
f1_keywords:
- vbawd10.chm158662713
ms.prod: word
api_name:
- Word.Selection.Cells
ms.assetid: 4b808b86-42ba-ccb4-b19a-87b134df3b79
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Cells property (Word)

Returns a  **[Cells](Word.cells.md)** collection that represents the table cells in a selection. Read-only.


## Syntax

_expression_.**Cells**

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example sets the current cell's background color to red.


```vb
If Selection.Information(wdWithInTable) = True Then 
 Selection.Cells(1).Shading.BackgroundPatternColorIndex = wdRed 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]