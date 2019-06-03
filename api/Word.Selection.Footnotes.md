---
title: Selection.Footnotes property (Word)
keywords: vbawd10.chm158662710
f1_keywords:
- vbawd10.chm158662710
ms.prod: word
api_name:
- Word.Selection.Footnotes
ms.assetid: 61829c93-46e9-c1c5-1424-fb34a812a76d
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Footnotes property (Word)

Returns a  **[Footnotes](Word.footnotes.md)** collection that represents all the footnotes in a range, selection, or document. Read-only.


## Syntax

_expression_. `Footnotes`

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example inserts an automatically numbered footnote at the insertion point.


```vb
Selection.Collapse Direction:=wdCollapseStart 
Selection.Footnotes.Add Range:=Selection.Range, _ 
 Text:="(Lone Creek Press, 1995)"
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]