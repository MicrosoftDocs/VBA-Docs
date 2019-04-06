---
title: AutoCorrect.ReplaceText property (Excel)
keywords: vbaxl10.chm545077
f1_keywords:
- vbaxl10.chm545077
ms.prod: excel
api_name:
- Excel.AutoCorrect.ReplaceText
ms.assetid: ff3321e3-335f-01a4-bbf2-2de8136d1d2d
ms.date: 04/06/2019
localization_priority: Normal
---


# AutoCorrect.ReplaceText property (Excel)

**True** if text in the list of AutoCorrect replacements is replaced automatically. Read/write **Boolean**.


## Syntax

_expression_.**ReplaceText**

_expression_ A variable that represents an **[AutoCorrect](Excel.AutoCorrect(object).md)** object.


## Example

This example turns off automatic text replacement.

```vb
With Application.AutoCorrect 
 .CapitalizeNamesOfDays = True 
 .ReplaceText = False 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]