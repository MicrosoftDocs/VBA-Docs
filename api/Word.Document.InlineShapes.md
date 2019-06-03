---
title: Document.InlineShapes property (Word)
keywords: vbawd10.chm158007364
f1_keywords:
- vbawd10.chm158007364
ms.prod: word
api_name:
- Word.Document.InlineShapes
ms.assetid: 049510b5-cdb3-74e8-783a-4c8fa809b876
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.InlineShapes property (Word)

Returns an  **[InlineShapes](Word.Document.InlineShapes.md)** collection that represents all the **[InlineShape](Word.InlineShape.md)** objects in a document. Read-only.


## Syntax

_expression_. `InlineShapes`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example displays the number of shapes and inline shapes in the active document.


```vb
Set doc = ActiveDocument 
Msgbox "InlineShape = " & doc.InlineShapes.Count & _ 
 vbCr & "Shapes = " & doc.Shapes.Count
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]