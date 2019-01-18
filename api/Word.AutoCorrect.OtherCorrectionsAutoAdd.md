---
title: AutoCorrect.OtherCorrectionsAutoAdd property (Word)
keywords: vbawd10.chm155779088
f1_keywords:
- vbawd10.chm155779088
ms.prod: word
api_name:
- Word.AutoCorrect.OtherCorrectionsAutoAdd
ms.assetid: ac284578-00af-7143-0573-a75a5557760c
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoCorrect.OtherCorrectionsAutoAdd property (Word)

 **True** if Microsoft Word automatically adds words to the list of AutoCorrect exceptions on the **Other Corrections** tab in the **AutoCorrect Exceptions** dialog box (**AutoCorrect Options** command, **Tools** menu). Word adds a word to this list if you delete and then retype a word that you didn't want Word to correct. Read/write **Boolean**.


## Syntax

 _expression_. `OtherCorrectionsAutoAdd`

 _expression_ An expression that returns an '[AutoCorrect](Word.AutoCorrect.md)' object.


## Example

This example sets Word to automatically add words to the list of AutoCorrect exceptions.


```vb
AutoCorrect.OtherCorrectionsAutoAdd = True
```


## See also


[AutoCorrect Object](Word.AutoCorrect.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]