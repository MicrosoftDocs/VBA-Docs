---
title: ErrorCheckingOptions.EmptyCellReferences property (Excel)
keywords: vbaxl10.chm698081
f1_keywords:
- vbaxl10.chm698081
api_name:
- Excel.ErrorCheckingOptions.EmptyCellReferences
ms.assetid: 3d9dd729-8483-aa8e-2d60-312bf3b3e08c
ms.date: 04/26/2019
ms.localizationpriority: medium
---


# ErrorCheckingOptions.EmptyCellReferences property (Excel)

When set to **True** (default), Microsoft Excel identifies, with an **AutoCorrect Options** button, selected cells containing formulas that refer to empty cells. **False** disables empty cell reference checking. Read/write **Boolean**.


## Syntax

_expression_.**EmptyCellReferences**

_expression_ A variable that represents an **[ErrorCheckingOptions](Excel.ErrorCheckingOptions.md)** object.


## Example

In the following example, the **AutoCorrect Options** button appears for cell A1, which contains a formula that references empty cells.

```vb
Sub CheckEmptyCells() 
 
 Application.ErrorCheckingOptions.EmptyCellReferences = True 
 Range("A1").Formula = "=A2+A3" 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]