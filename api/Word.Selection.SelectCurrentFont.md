---
title: Selection.SelectCurrentFont method (Word)
keywords: vbawd10.chm158663173
f1_keywords:
- vbawd10.chm158663173
ms.prod: word
api_name:
- Word.Selection.SelectCurrentFont
ms.assetid: 66539ab3-280f-40a5-1fc0-1507b66d50fd
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.SelectCurrentFont method (Word)

Extends the selection forward until text in a different font or font size is encountered.


## Syntax

_expression_. `SelectCurrentFont`

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Example

This example extends the selection until text in a different font or font size is encountered. The example uses the **Grow** method to increase the size of the selected text to the next available font size.


```vb
With Selection 
 .SelectCurrentFont 
 .Font.Grow 
End With
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]