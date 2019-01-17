---
title: AutoCorrect.CorrectSentenceCap property (Excel)
keywords: vbaxl10.chm545079
f1_keywords:
- vbaxl10.chm545079
ms.prod: excel
api_name:
- Excel.AutoCorrect.CorrectSentenceCap
ms.assetid: 3e37a79e-8199-4bd1-3601-f51243782128
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoCorrect.CorrectSentenceCap property (Excel)

 **True** if Microsoft Excel automatically corrects sentence (first word) capitalization. Read/write **Boolean**.


## Syntax

_expression_. `CorrectSentenceCap`

_expression_ A variable that represents an [AutoCorrect](Excel.AutoCorrect-graph-property.md) object.


## Example

This example enables Microsoft Excel to automatically correct sentence capitalization.


```vb
Application.AutoCorrect.CorrectSentenceCap = True
```


## See also


[AutoCorrect Object](Excel.AutoCorrect(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]