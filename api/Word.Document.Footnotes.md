---
title: Document.Footnotes property (Word)
keywords: vbawd10.chm158007303
f1_keywords:
- vbawd10.chm158007303
ms.prod: word
api_name:
- Word.Document.Footnotes
ms.assetid: 6257f658-69f5-4223-153b-56bc3791a99d
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Footnotes property (Word)

Returns a  **[Footnotes](Word.footnotes.md)** collection that represents all the footnotes in a document. Read-only.


## Syntax

_expression_. `Footnotes`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example changes the footnote reference marks for the footnotes in the active document to lowercase letters, starting with the letter "c".


```vb
With ActiveDocument.Footnotes 
 .StartingNumber = 3 
 .NumberStyle = wdNoteNumberStyleLowercaseLetter 
End With
```

This example inserts an automatically numbered footnote at the insertion point.




```vb
Selection.Collapse Direction:=wdCollapseStart 
ActiveDocument.Footnotes.Add Range:=Selection.Range, _ 
 Text:="(Lone Creek Press, 1995)"
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]