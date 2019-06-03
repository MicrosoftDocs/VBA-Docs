---
title: Document.Sentences property (Word)
keywords: vbawd10.chm158007314
f1_keywords:
- vbawd10.chm158007314
ms.prod: word
api_name:
- Word.Document.Sentences
ms.assetid: 41906136-815c-4dfc-ad92-c16ad420ab91
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Sentences property (Word)

Returns a  **[Sentences](Word.sentences.md)** collection that represents all the sentences in the document. Read-only.


## Syntax

_expression_. `Sentences`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example copies the first sentences in the active document.


```vb
ActiveDocument.Sentences(1).Copy
```

This example deletes the last sentence in the active document.




```vb
ActiveDocument.Sentences.Last.Delete
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]