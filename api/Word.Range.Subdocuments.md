---
title: Range.Subdocuments property (Word)
keywords: vbawd10.chm157155487
f1_keywords:
- vbawd10.chm157155487
ms.prod: word
api_name:
- Word.Range.Subdocuments
ms.assetid: c06afeb9-7e83-d858-d863-9582962c8254
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.Subdocuments property (Word)

Returns a  **Subdocuments** collection that represents all the subdocuments in the specified range or document. Read-only.


## Syntax

_expression_. `Subdocuments`

_expression_ A variable that represents a **[Range](Word.Range.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example displays the number of subdocuments embedded in the active document.


```vb
MsgBox ActiveDocument.Range.Subdocuments.Count
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]