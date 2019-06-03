---
title: Document.TablesOfFigures property (Word)
keywords: vbawd10.chm158007321
f1_keywords:
- vbawd10.chm158007321
ms.prod: word
api_name:
- Word.Document.TablesOfFigures
ms.assetid: 1c386611-82f9-0a0d-71ce-dfe006d8eab5
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.TablesOfFigures property (Word)

Returns a  **[TablesOfFigures](Word.Document.TablesOfFigures.md)** collection that represents the tables of figures in the specified document. Read-only.


## Syntax

_expression_. `TablesOfFigures`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example adds a table of figures at the insertion point in the active document.


```vb
Selection.Collapse Direction:=wdCollapseStart 
ActiveDocument.TablesOfFigures.Add Range:=Selection.Range, _ 
 Caption:=wdCaptionFigure
```

This example updates the contents of the first table of figures in Report.doc.




```vb
Documents("Report.doc").TablesOfFigures(1).Update
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]