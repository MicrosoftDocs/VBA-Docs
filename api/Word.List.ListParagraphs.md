---
title: List.ListParagraphs property (Word)
keywords: vbawd10.chm160563202
f1_keywords:
- vbawd10.chm160563202
ms.prod: word
api_name:
- Word.List.ListParagraphs
ms.assetid: 3360f8dd-155a-3b44-1b0c-395ddbac2b51
ms.date: 06/08/2017
localization_priority: Normal
---


# List.ListParagraphs property (Word)

Returns a  **[ListParagraphs](Word.listparagraphs.md)** collection that represents all the numbered paragraphs in the list, document, or range. Read-only.


## Syntax

_expression_. `ListParagraphs`

_expression_ A variable that represents a '[List](Word.List.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example double underlines the paragraphs in the second list in the active document.


```vb
For Each mypara In ActiveDocument.Lists(2).ListParagraphs 
 mypara.Range.Underline = wdUnderlineDouble 
Next mypara
```


## See also


[List Object](Word.List.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]