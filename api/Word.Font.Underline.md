---
title: Font.Underline property (Word)
keywords: vbawd10.chm156369036
f1_keywords:
- vbawd10.chm156369036
ms.prod: word
api_name:
- Word.Font.Underline
ms.assetid: 3fbcecb6-c38c-746e-671a-1339aa855b15
ms.date: 06/08/2017
localization_priority: Normal
---


# Font.Underline property (Word)

Returns or sets the type of underline applied to the font. Read/write  **[WdUnderline](Word.WdUnderline.md)**.


## Syntax

 _expression_. `Underline`

 _expression_ Required. A variable that represents a '[Font](Word.Font.md)' object.


## Example

This example applies a single underline to the selected text.


```vb
If Selection.Type = wdSelectionNormal Then 
 Selection.Font.Underline = wdUnderlineSingle 
Else 
 MsgBox "You need to select some text." 
End If
```


## See also


[Font Object](Word.Font.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]