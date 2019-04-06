---
title: AutoCorrect.CapitalizeNamesOfDays property (Excel)
keywords: vbaxl10.chm545074
f1_keywords:
- vbaxl10.chm545074
ms.prod: excel
api_name:
- Excel.AutoCorrect.CapitalizeNamesOfDays
ms.assetid: 218f9820-8cb1-438d-7c81-4a9c4385a275
ms.date: 04/06/2019
localization_priority: Normal
---


# AutoCorrect.CapitalizeNamesOfDays property (Excel)

**True** if the first letter of day names is capitalized automatically. Read/write **Boolean**.


## Syntax

_expression_.**CapitalizeNamesOfDays**

_expression_ A variable that represents an **[AutoCorrect](Excel.AutoCorrect(object).md)** object.


## Example

This example sets Microsoft Excel to capitalize the first letter of the names of days.

```vb
With Application.AutoCorrect 
 .CapitalizeNamesOfDays = True 
 .ReplaceText = True 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]