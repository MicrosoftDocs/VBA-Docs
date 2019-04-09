---
title: Template.JustificationMode property (Word)
keywords: vbawd10.chm157941773
f1_keywords:
- vbawd10.chm157941773
ms.prod: word
api_name:
- Word.Template.JustificationMode
ms.assetid: 914994e8-8ea3-4119-271c-193970da060c
ms.date: 06/08/2017
localization_priority: Normal
---


# Template.JustificationMode property (Word)

Returns or sets the character spacing adjustment for the specified template. Read/write  **[WdJustificationMode](Word.WdJustificationMode.md)**.


## Syntax

_expression_. `JustificationMode`

_expression_ Required. A variable that represents a '[Template](Word.Template.md)' object.


## Example

This example sets Microsoft Word to compress only punctuation marks when adjusting character spacing.


```vb
NormalTemplate.JustificationMode = wdJustificationModeCompressKana
```


## See also


[Template Object](Word.Template.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]