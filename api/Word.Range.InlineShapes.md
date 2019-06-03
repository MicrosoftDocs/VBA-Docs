---
title: Range.InlineShapes property (Word)
keywords: vbawd10.chm157155647
f1_keywords:
- vbawd10.chm157155647
ms.prod: word
api_name:
- Word.Range.InlineShapes
ms.assetid: 4c0335ac-95a2-412c-650c-afc323ae58ca
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.InlineShapes property (Word)

Returns an  **InlineShapes** collection that represents all the **InlineShape** objects in a range. Read-only.


## Syntax

_expression_. `InlineShapes`

_expression_ A variable that represents a **[Range](Word.Range.md)** object.


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


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]