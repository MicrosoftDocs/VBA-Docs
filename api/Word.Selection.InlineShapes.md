---
title: Selection.InlineShapes property (Word)
keywords: vbawd10.chm158663067
f1_keywords:
- vbawd10.chm158663067
ms.prod: word
api_name:
- Word.Selection.InlineShapes
ms.assetid: 2fbbf39c-b70e-e332-2547-089166e718ca
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.InlineShapes property (Word)

Returns an  **[InlineShapes](Word.inlineshapes.md)** collection that represents all the **InlineShape** objects in a selection. Read-only.


## Syntax

_expression_. `InlineShapes`

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


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


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]