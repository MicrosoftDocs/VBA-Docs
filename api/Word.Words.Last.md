---
title: Words.Last property (Word)
keywords: vbawd10.chm157024260
f1_keywords:
- vbawd10.chm157024260
ms.prod: word
api_name:
- Word.Words.Last
ms.assetid: 5ca384f7-786f-9c44-41fb-4dce72d45d3e
ms.date: 06/08/2017
localization_priority: Normal
---


# Words.Last property (Word)

Returns a  **Range** object that represents the last word in a collection of words.


## Syntax

 _expression_. `Last`

 _expression_ Required. A variable that represents a '[Words](Word.words.md)' collection.


## Example

This example applies bold formatting to the last word in the selection.


```vb
If Selection.Words.Count >= 2 Then 
 Selection.Words.Last.Bold = True 
End If
```


## See also


[Words Collection Object](Word.words.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]