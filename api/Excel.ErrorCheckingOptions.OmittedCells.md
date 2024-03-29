---
title: ErrorCheckingOptions.OmittedCells property (Excel)
keywords: vbaxl10.chm698079
f1_keywords:
- vbaxl10.chm698079
api_name:
- Excel.ErrorCheckingOptions.OmittedCells
ms.assetid: a337da5d-4f02-d24c-c59a-288b4a9c9117
ms.date: 04/26/2019
ms.localizationpriority: medium
---


# ErrorCheckingOptions.OmittedCells property (Excel)

When set to **True** (default), Microsoft Excel identifies, with an **AutoCorrect Options** button, the selected cells that contain formulas referring to a range that omits adjacent cells that could be included. **False** disables error checking for omitted cells. Read/write **Boolean**.


## Syntax

_expression_.**OmittedCells**

_expression_ A variable that represents an **[ErrorCheckingOptions](Excel.ErrorCheckingOptions.md)** object.


## Example

In the following example, the **AutoCorrect Options** button appears for cell A4, which contains a formula.

```vb
Sub CheckOmittedCells() 
 
 Application.ErrorCheckingOptions.OmittedCells = True 
 Range("A1").Value = 1 
 Range("A2").Value = 2 
 Range("A3").Value = 3 
 Range("A4").Formula = "=Sum(A1:A2)" 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]