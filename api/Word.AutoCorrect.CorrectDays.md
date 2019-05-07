---
title: AutoCorrect.CorrectDays property (Word)
keywords: vbawd10.chm155779073
f1_keywords:
- vbawd10.chm155779073
ms.prod: word
api_name:
- Word.AutoCorrect.CorrectDays
ms.assetid: a9b4ee11-72bf-41d7-883f-6cacd13ed770
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoCorrect.CorrectDays property (Word)

 **True** if Word automatically capitalizes the first letter of days of the week. Read/write **Boolean**.


## Syntax

_expression_. `CorrectDays`

_expression_ A variable that represents an '[AutoCorrect](Word.AutoCorrect.md)' object.


## Example

This example sets Word to automatically capitalize the first letter of days of the week.


```vb
AutoCorrect.CorrectDays = True
```

This example toggles the value of the CorrectDays property.




```vb
AutoCorrect.CorrectDays = Not AutoCorrect.CorrectDays
```


## See also


[AutoCorrect Object](Word.AutoCorrect.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]