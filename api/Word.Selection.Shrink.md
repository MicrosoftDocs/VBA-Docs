---
title: Selection.Shrink method (Word)
keywords: vbawd10.chm158662957
f1_keywords:
- vbawd10.chm158662957
ms.prod: word
api_name:
- Word.Selection.Shrink
ms.assetid: ed364c95-3b9d-44dc-b120-db23aedfeaed
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Shrink method (Word)

Shrinks the selection to the next smaller unit of text.


## Syntax

_expression_. `Shrink`

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

The unit progression for this method is as follows: entire document, section, paragraph, sentence, word, insertion point.


## Example

This example collapses the selected text to the next smaller unit of text.


```vb
If Selection.Type = wdSelectionNormal Then 
 Selection.Shrink 
Else 
 MsgBox "You need to select some text." 
End If
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]