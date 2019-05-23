---
title: Range.Select method (Word)
keywords: vbawd10.chm157220863
f1_keywords:
- vbawd10.chm157220863
ms.prod: word
api_name:
- Word.Range.Select
ms.assetid: 732c2aca-d8b4-3537-984f-d44d4eed870a
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.Select method (Word)

Selects the specified range.


## Syntax

_expression_.**Select**

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Example

This example selects the first paragraph in the active document.


```vb
Sub SelectParagraph() 
 ActiveDocument.Paragraphs(1).Range.Select 
 Selection.Font.Bold = True 
End Sub
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
