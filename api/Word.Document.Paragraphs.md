---
title: Document.Paragraphs property (Word)
keywords: vbawd10.chm158007312
f1_keywords:
- vbawd10.chm158007312
ms.prod: word
api_name:
- Word.Document.Paragraphs
ms.assetid: ad60de6b-6287-8ea0-142e-8795f623aa29
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Paragraphs property (Word)

Returns a  **Paragraphs** collection that represents all the paragraphs in the specified document. Read-only.


## Syntax

_expression_. `Paragraphs`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example sets the line spacing to single for the collection of all paragraphs in section one in the active document.


```vb
ActiveDocument.Sections(1).Range.Paragraphs.LineSpacingRule = _ 
 wdLineSpaceSingle
```

This example sets the line spacing to double for the first paragraph in the selection.




```vb
Selection.Paragraphs(1).LineSpacingRule = wdLineSpaceDouble
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]