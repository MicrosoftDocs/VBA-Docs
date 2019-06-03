---
title: Document.Frames property (Word)
keywords: vbawd10.chm158007319
f1_keywords:
- vbawd10.chm158007319
ms.prod: word
api_name:
- Word.Document.Frames
ms.assetid: 61b7d5dc-6ab4-d29c-6c6e-daac6a2431ed
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Frames property (Word)

Returns a  **[Frames](Word.Frames.md)** collection that represents all the frames in a document. Read-only.


## Syntax

_expression_. `Frames`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example adds a frame around the selection and returns a frame object to the myFrame variable.


```vb
Set myFrame = ActiveDocument.Frames.Add(Range:=Selection.Range)
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]