---
title: Find.CorrectHangulEndings property (Word)
keywords: vbawd10.chm162529341
f1_keywords:
- vbawd10.chm162529341
ms.prod: word
api_name:
- Word.Find.CorrectHangulEndings
ms.assetid: 814affac-ba96-7e93-6c58-6d063c15b79c
ms.date: 06/08/2017
localization_priority: Normal
---


# Find.CorrectHangulEndings property (Word)

 **True** if Microsoft Word automatically corrects Hangul endings when replacing Hangul text. Read/write **Boolean**.


## Syntax

_expression_. `CorrectHangulEndings`

 _expression_ An expression that returns a '[Find](Word.Find.md)' object.


## Example

This example sets Microsoft Word to automatically correct Hangul endings when replacing Hangul text.


```vb
With Selection.Find 
 .Forward = True 
 .Wrap = wdFindContinue 
 .Format = False 
 .CorrectHangulEndings = True 
End With
```


## See also


[Find Object](Word.Find.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]