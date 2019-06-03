---
title: Selection.Paragraphs property (Word)
keywords: vbawd10.chm158662715
f1_keywords:
- vbawd10.chm158662715
ms.prod: word
api_name:
- Word.Selection.Paragraphs
ms.assetid: f237788a-01e4-62ce-d698-3af619c90272
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Paragraphs property (Word)

Returns a  **[Paragraphs](Word.paragraphs.md)** collection that represents all the paragraphs in the specified selection. Read-only.


## Syntax

_expression_. `Paragraphs`

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example sets the line spacing to double for the first paragraph in the selection.


```vb
Selection.Paragraphs(1).LineSpacingRule = wdLineSpaceDouble
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]