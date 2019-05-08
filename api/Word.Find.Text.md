---
title: Find.Text property (Word)
keywords: vbawd10.chm162529302
f1_keywords:
- vbawd10.chm162529302
ms.prod: word
api_name:
- Word.Find.Text
ms.assetid: d92917aa-32f7-e9cc-bb74-03f7ed17498a
ms.date: 06/08/2017
localization_priority: Normal
---


# Find.Text property (Word)

Returns or sets the text to find. Read/write  **String**.


## Syntax

_expression_.**Text**

_expression_ A variable that represents a '[Find](Word.Find.md)' object.


## Remarks

The  **Text** property returns the plain, unformatted text of the selection or range. When you set this property, the text of the range or selection is replaced.


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


[Find Object](Word.Find.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]