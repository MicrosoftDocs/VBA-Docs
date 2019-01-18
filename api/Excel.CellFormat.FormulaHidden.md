---
title: CellFormat.FormulaHidden property (Excel)
keywords: vbaxl10.chm676086
f1_keywords:
- vbaxl10.chm676086
ms.prod: excel
api_name:
- Excel.CellFormat.FormulaHidden
ms.assetid: 5e1b6875-f66a-568a-e0e5-af88e64edfe6
ms.date: 06/08/2017
localization_priority: Normal
---


# CellFormat.FormulaHidden property (Excel)

Returns or sets a  **Variant** value that indicates if the formula will be hidden when the worksheet is protected.


## Syntax

_expression_. `FormulaHidden`

_expression_ A variable that represents a [CellFormat](Excel.CellFormat.md) object.


## Remarks

This property returns  **True** if the formula will be hidden when the worksheet is protected, **Null** if the specified range contains some cells with **FormulaHidden** equal to **True** and some cells with **FormulaHidden** equal to **False**.

Don't confuse this property with the  **[Hidden](Excel.Range.Hidden.md)** property. The formula will not be hidden if the workbook is protected and the worksheet is not, but only if the worksheet is protected.


## See also


[CellFormat Object](Excel.CellFormat.md)

