---
title: PageSetup.TextColumns property (Word)
keywords: vbawd10.chm158400631
f1_keywords:
- vbawd10.chm158400631
ms.prod: word
api_name:
- Word.PageSetup.TextColumns
ms.assetid: 85fb6b50-1c2e-a96e-e56d-3a1caacaabc5
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup.TextColumns property (Word)

Returns a  **[TextColumns](Word(textcolumns).md)** collection that represents the set of text columns for the specified **PageSetup** object.


## Syntax

_expression_. `TextColumns`

 _expression_ An expression that returns a **[PageSetup](Word.PageSetup.md)** object.


## Remarks

There will always be at least one text column in the collection. When you create new text columns, you are adding to a collection of one column.

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example creates four evenly-spaced text columns that are applied to section two in the active document.


```vb
With ActiveDocument.Sections(2).PageSetup.TextColumns 
 .SetCount NumColumns:=3 
 .Add EvenlySpaced:=True 
End With
```

This example creates a document with two text columns. The first column is 1.5 inches wide and the second is 3 inches wide.




```vb
Set myDoc = Documents.Add 
With myDoc.PageSetup.TextColumns 
 .SetCount NumColumns:=1 
 .Add Width:=InchesToPoints(3) 
End With 
With myDoc.PageSetup.TextColumns(1) 
 .Width = InchesToPoints(1.5) 
 .SpaceAfter = InchesToPoints(0.5) 
End With
```


## See also


[PageSetup Object](Word.PageSetup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]