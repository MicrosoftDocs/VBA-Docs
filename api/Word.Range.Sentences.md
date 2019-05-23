---
title: Range.Sentences property (Word)
keywords: vbawd10.chm157155380
f1_keywords:
- vbawd10.chm157155380
ms.prod: word
api_name:
- Word.Range.Sentences
ms.assetid: fe870f13-d09f-efbf-1d2f-745f2c318c28
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.Sentences property (Word)

Returns a  **Sentences** collection that represents all the sentences in the range. Read-only.


## Syntax

_expression_. `Sentences`

_expression_ A variable that represents a **[Range](Word.Range.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning a Single Object from a Collection](../word/Concepts/Miscellaneous/returning-a-single-object-from-a-collection.md).


## Example

This example displays the number of sentences in the first paragraph in the active document.


```vb
MsgBox ActiveDocument.Paragraphs(1).Range _ 
 .Sentences.Count & " sentences"
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]