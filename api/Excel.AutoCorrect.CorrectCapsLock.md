---
title: AutoCorrect.CorrectCapsLock property (Excel)
keywords: vbaxl10.chm545080
f1_keywords:
- vbaxl10.chm545080
ms.prod: excel
api_name:
- Excel.AutoCorrect.CorrectCapsLock
ms.assetid: 02a1944c-03fb-3727-b2d3-9da04f7e74a4
ms.date: 04/06/2019
localization_priority: Normal
---


# AutoCorrect.CorrectCapsLock property (Excel)

**True** if Microsoft Excel automatically corrects accidental use of the CapsLock key. Read/write **Boolean**.


## Syntax

_expression_.**CorrectCapsLock**

_expression_ A variable that represents an **[AutoCorrect](Excel.AutoCorrect(object).md)** object.


## Example

This example enables Excel to automatically correct accidental use of the CapsLock key.

```vb
Application.AutoCorrect.CorrectCapsLock = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]