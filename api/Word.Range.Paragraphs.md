---
title: Range.Paragraphs property (Word)
keywords: vbawd10.chm157155387
f1_keywords:
- vbawd10.chm157155387
ms.prod: word
api_name:
- Word.Range.Paragraphs
ms.assetid: b5c9df62-a477-ce1a-4a94-027100527a6f
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.Paragraphs property (Word)

Returns a  **Paragraphs** collection that represents all the paragraphs in the specified range. Read-only.


## Syntax

_expression_. `Paragraphs`

_expression_ A variable that represents a **[Range](Word.Range.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example sets the line spacing to single for the collection of all paragraphs in section one in the active document.


```vb
ActiveDocument.Sections(1).Range.Paragraphs.LineSpacingRule = _ 
 wdLineSpaceSingle
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]