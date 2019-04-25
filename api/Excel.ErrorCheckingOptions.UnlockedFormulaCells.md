---
title: ErrorCheckingOptions.UnlockedFormulaCells property (Excel)
keywords: vbaxl10.chm698080
f1_keywords:
- vbaxl10.chm698080
ms.prod: excel
api_name:
- Excel.ErrorCheckingOptions.UnlockedFormulaCells
ms.assetid: 0b7c038d-41d8-aeb8-3e15-3105d6e65c02
ms.date: 04/26/2019
localization_priority: Normal
---


# ErrorCheckingOptions.UnlockedFormulaCells property (Excel)

When set to **True** (default), Microsoft Excel identifies selected cells that are unlocked and contain a formula. **False** disables error checking for unlocked cells that contain formulas. Read/write **Boolean**.


## Syntax

_expression_.**UnlockedFormulaCells**

_expression_ A variable that represents an **[ErrorCheckingOptions](Excel.ErrorCheckingOptions.md)** object.


## Example

In the following example, the **AutoCorrect Options** button appears for cell A3, which is an unlocked cell containing a formula.

```vb
Sub CheckUnlockedCell() 
 
 Application.ErrorCheckingOptions.UnlockedFormulaCells = True 
 Range("A1").Value = 1 
 Range("A2").Value = 2 
 Range("A3").Formula = "=A1+A2" 
 Range("A3").Locked = False 
 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]