---
title: AutoCorrectEntry.RichText property (Word)
keywords: vbawd10.chm155648004
f1_keywords:
- vbawd10.chm155648004
ms.prod: word
api_name:
- Word.AutoCorrectEntry.RichText
ms.assetid: f612473f-d051-1b22-3274-dbd0dd8c49ac
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoCorrectEntry.RichText property (Word)

 **True** if formatting is stored with the AutoCorrect entry replacement text. Read-only **Boolean**.


## Syntax

 _expression_. `RichText`

 _expression_ An expression that returns an '[AutoCorrectEntry](Word.AutoCorrectEntry.md)' object.


## Example

This example determines whether AutoCorrect entry one is formatted.


```vb
MsgBox AutoCorrect.Entries(1).RichText
```


## See also


[AutoCorrectEntry Object](Word.AutoCorrectEntry.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]