---
title: Document.Endnotes property (Word)
keywords: vbawd10.chm158007304
f1_keywords:
- vbawd10.chm158007304
ms.prod: word
api_name:
- Word.Document.Endnotes
ms.assetid: 3c3e87c0-ea76-8bc4-0b2e-755bff6aa14c
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Endnotes property (Word)

Returns an  **[Endnotes](Word.endnotes.md)** collection that represents all the endnotes in a document. Read-only.


## Syntax

_expression_. `Endnotes`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example positions the endnotes in the active document at the end of the document and formats the endnote reference marks as lowercase roman numerals.


```vb
With ActiveDocument.Endnotes 
 .Location = wdEndOfDocument 
 .NumberStyle = wdNoteNumberStyleLowercaseRoman 
End With
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]