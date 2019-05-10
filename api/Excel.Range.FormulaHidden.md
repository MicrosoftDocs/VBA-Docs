---
title: Range.FormulaHidden property (Excel)
keywords: vbaxl10.chm144135
f1_keywords:
- vbaxl10.chm144135
ms.prod: excel
api_name:
- Excel.Range.FormulaHidden
ms.assetid: b6425c86-7e20-e34e-2d96-eb16075c20b6
ms.date: 05/10/2019
localization_priority: Normal
---


# Range.FormulaHidden property (Excel)

Returns or sets a **Variant** value that indicates if the formula will be hidden when the worksheet is protected.


## Syntax

_expression_.**FormulaHidden**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

This property returns **True** if the formula will be hidden when the worksheet is protected, **Null** if the specified range contains some cells with **FormulaHidden** equal to **True** and some cells with **FormulaHidden** equal to **False**.

Don't confuse this property with the **[Hidden](Excel.Range.Hidden.md)** property. The formula will not be hidden if the workbook is protected and the worksheet is not, but only if the worksheet is protected.


## Example

This example hides the formulas in cells A1 and B1 on Sheet1 when the worksheet is protected.

```vb
Sub HideFormulas() 
 
 Worksheets("Sheet1").Range("A1:B1").FormulaHidden = True 
 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
