---
title: Selection.Words property (Word)
keywords: vbawd10.chm158662707
f1_keywords:
- vbawd10.chm158662707
ms.prod: word
api_name:
- Word.Selection.Words
ms.assetid: bbbc7c5f-ce5a-2608-ba0c-e9769bff287a
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Words property (Word)

Returns a  **[Words](Word.words.md)** collection that represents all the words in a selection. Read-only.


## Syntax

_expression_. `Words`

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

Punctuation and paragraph marks in a document are included in the  **[Words](Word.words.md)** collection. For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example displays the number of words in the selection. Paragraphs marks, partial words, and punctuation are included in the count.


```vb
MsgBox "There are " & Selection.Words.Count & " words."
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]