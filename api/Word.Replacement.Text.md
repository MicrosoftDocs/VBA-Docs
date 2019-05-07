---
title: Replacement.Text property (Word)
keywords: vbawd10.chm162594831
f1_keywords:
- vbawd10.chm162594831
ms.prod: word
api_name:
- Word.Replacement.Text
ms.assetid: bfd99129-8d38-b448-6d74-cda1a78304d7
ms.date: 06/08/2017
localization_priority: Normal
---


# Replacement.Text property (Word)

Returns or sets the text to replace. Read/write  **String**.


## Syntax

_expression_.**Text**

_expression_ A variable that represents a '[Replacement](Word.Replacement.md)' object.


## Example

This example replaces "Hello" with "Goodbye" in the active document.


```vb
Set myRange = ActiveDocument.Content 
With myRange.Find 
 .ClearFormatting 
 .Replacement.ClearFormatting 
 .Text = "Hello" 
 .Replacement.Text = "Goodbye" 
 .Execute Replace:=wdReplaceAll 
End With
```


## See also


[Replacement Object](Word.Replacement.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]