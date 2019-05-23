---
title: Range.Revisions property (Word)
keywords: vbawd10.chm157155478
f1_keywords:
- vbawd10.chm157155478
ms.prod: word
api_name:
- Word.Range.Revisions
ms.assetid: cf71b684-991a-fb6d-09bc-eeecb16edec5
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.Revisions property (Word)

Returns a  **Revisions** collection that represents the tracked changes in the range. Read-only.


## Syntax

_expression_. `Revisions`

_expression_ A variable that represents a **[Range](Word.Range.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning a Single Object from a Collection](../word/Concepts/Miscellaneous/returning-a-single-object-from-a-collection.md).


## Example

This example displays the number of tracked changes in the first section in the active document.


```vb
MsgBox ActiveDocument.Sections(1).Range.Revisions.Count
```

This example accepts all tracked changes in the first paragraph in the selection.




```vb
Set myRange = Selection.Paragraphs(1).Range 
myRange.Revisions.AcceptAll
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]