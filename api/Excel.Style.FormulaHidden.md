---
title: Style.FormulaHidden property (Excel)
keywords: vbaxl10.chm177078
f1_keywords:
- vbaxl10.chm177078
ms.prod: excel
api_name:
- Excel.Style.FormulaHidden
ms.assetid: 7b36f86b-2f88-3fb4-173e-cca7e747a195
ms.date: 06/08/2017
localization_priority: Normal
---


# Style.FormulaHidden property (Excel)

Returns or sets a  **Boolean** value that indicates if the formula will be hidden when the worksheet is protected.


## Syntax

_expression_. `FormulaHidden`

_expression_ A variable that represents a [Style](Excel.Style.md) object.


## Remarks

Don't confuse this property with the  **[Hidden](Excel.Range.Hidden.md)** property. The formula will not be hidden if the workbook is protected and the worksheet is not, but only if the worksheet is protected.


## See also


[Style Object](Excel.Style.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]