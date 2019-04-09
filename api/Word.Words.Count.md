---
title: Words.Count property (Word)
keywords: vbawd10.chm157024258
f1_keywords:
- vbawd10.chm157024258
ms.prod: word
api_name:
- Word.Words.Count
ms.assetid: abbb4293-0ffb-f845-cdda-acbbe0ff477b
ms.date: 06/08/2017
localization_priority: Normal
---


# Words.Count property (Word)

Returns a  **Long** that represents the number of words in the collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ Required. A variable that represents a '[Words](Word.words.md)' collection.


## Example

This example displays the number of words in the selection.


```vb
If Selection.Words.Count >= 1 And _ 
 Selection.Type <> wdSelectionIP Then 
 MsgBox "The selection contains " & Selection.Words.Count _ 
 & " words." 
End If
```


## See also


[Words Collection Object](Word.words.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]