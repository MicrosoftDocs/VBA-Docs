---
title: Document.Characters property (Word)
keywords: vbawd10.chm158007315
f1_keywords:
- vbawd10.chm158007315
ms.prod: word
api_name:
- Word.Document.Characters
ms.assetid: 1703bbe3-6c46-a45b-9f36-1205a0d2d47c
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Characters property (Word)

Returns a  **[Characters](Word.characters.md)** collection that represents the characters in a document. Read-only.


## Syntax

_expression_. `Characters`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example returns the number of characters in the first sentence in the active document (spaces are included in the count).


```vb
numchars = ActiveDocument.Characters.Count
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]