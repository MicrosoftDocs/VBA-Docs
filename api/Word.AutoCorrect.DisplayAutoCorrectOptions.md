---
title: AutoCorrect.DisplayAutoCorrectOptions property (Word)
keywords: vbawd10.chm155779092
f1_keywords:
- vbawd10.chm155779092
ms.prod: word
api_name:
- Word.AutoCorrect.DisplayAutoCorrectOptions
ms.assetid: 7a4d6773-53f7-8d9d-499e-8d32917c14fd
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoCorrect.DisplayAutoCorrectOptions property (Word)

 **True** for Microsoft Word to display the **AutoCorrect Options** button. Read/write **Boolean**.


## Syntax

_expression_. `DisplayAutoCorrectOptions`

 _expression_ An expression that returns an '[AutoCorrect](Word.AutoCorrect.md)' object.


## Example

This example disables display of the **AutoCorrect Options** button.


```vb
Sub HideAutoCorrectOpButton() 
 AutoCorrect.DisplayAutoCorrectOptions = False 
End Sub
```


## See also


[AutoCorrect Object](Word.AutoCorrect.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]