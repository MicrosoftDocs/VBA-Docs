---
title: Find.Font property (Word)
keywords: vbawd10.chm162529291
f1_keywords:
- vbawd10.chm162529291
ms.prod: word
api_name:
- Word.Find.Font
ms.assetid: 8a4e3cb0-5bfd-bcea-6eba-10dc21a0e4c0
ms.date: 06/08/2017
localization_priority: Normal
---


# Find.Font property (Word)

Returns or sets a  **[Font](Word.Font.md)** object that represents the character formatting of the specified object. Read/write **Font**.


## Syntax

_expression_.**Font**

_expression_ A variable that represents a '[Find](Word.Find.md)' object.


## Remarks

To set this property, specify an expression that returns a  **[Font](Word.Font.md)** object.


## Example

This example finds the next range of text that's formatted with the Times New Roman font.


```vb
With Selection.Find 
 .ClearFormatting 
 .Font.Name = "Times New Roman" 
 .Execute FindText:="", ReplaceWith:="", Format:=True, _ 
 Forward:=True 
End With
```


## See also


[Find Object](Word.Find.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]