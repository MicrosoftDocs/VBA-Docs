---
title: Font.EmphasisMark property (Word)
keywords: vbawd10.chm156369050
f1_keywords:
- vbawd10.chm156369050
ms.prod: word
api_name:
- Word.Font.EmphasisMark
ms.assetid: 18e541c3-09aa-690d-94fa-ace6133c5cc6
ms.date: 06/08/2017
localization_priority: Normal
---


# Font.EmphasisMark property (Word)

Returns or sets a  **WdEmphasisMark** constant that represents the emphasis mark for a character or designated character string. Read/write.


## Syntax

 _expression_. `EmphasisMark`

 _expression_ Required. A variable that represents a '[Font](Word.Font.md)' object.


## Example

This example sets the emphasis mark over the fourth word in the active document to a comma.


```vb
ActiveDocument.Words(4).EmphasisMark = wdEmphasisMarkOverComma
```


## See also


[Font Object](Word.Font.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]