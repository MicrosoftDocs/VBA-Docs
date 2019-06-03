---
title: Selection.Endnotes property (Word)
keywords: vbawd10.chm158662711
f1_keywords:
- vbawd10.chm158662711
ms.prod: word
api_name:
- Word.Selection.Endnotes
ms.assetid: fea9ea39-4091-cccd-9025-36be2e4b95ff
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Endnotes property (Word)

Returns an  **[Endnotes](Word.endnotes.md)** collection that represents all the endnotes contained within a selection. Read-only.


## Syntax

_expression_. `Endnotes`

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example positions the endnotes in the selection at the end of the document and formats the endnote reference marks as lowercase roman numerals.


```vb
With Selection.Endnotes 
 .Location = wdEndOfDocument 
 .NumberStyle = wdNoteNumberStyleLowercaseRoman 
End With
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]