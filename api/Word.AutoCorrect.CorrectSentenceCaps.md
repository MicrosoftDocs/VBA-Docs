---
title: AutoCorrect.CorrectSentenceCaps property (Word)
keywords: vbawd10.chm155779075
f1_keywords:
- vbawd10.chm155779075
ms.prod: word
api_name:
- Word.AutoCorrect.CorrectSentenceCaps
ms.assetid: 47eb861a-2dcc-27c9-33ee-5e5bc0d6df4b
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoCorrect.CorrectSentenceCaps property (Word)

 **True** if Word automatically capitalizes the first letter in each sentence. Read/write **Boolean**.


## Syntax

_expression_. `CorrectSentenceCaps`

_expression_ A variable that represents an '[AutoCorrect](Word.AutoCorrect.md)' object.


## Example

This example toggles the value of the CorrectSentenceCaps property.


```vb
AutoCorrect.CorrectSentenceCaps = Not _ 
 AutoCorrect.CorrectSentenceCaps
```


## See also


[AutoCorrect Object](Word.AutoCorrect.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]