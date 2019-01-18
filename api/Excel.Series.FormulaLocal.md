---
title: Series.FormulaLocal property (Excel)
keywords: vbaxl10.chm578085
f1_keywords:
- vbaxl10.chm578085
ms.prod: excel
api_name:
- Excel.Series.FormulaLocal
ms.assetid: 6e2a0912-5006-d223-30a6-618642de035d
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.FormulaLocal property (Excel)

Returns or sets the formula for the object, using A1-style references in the language of the user. Read/write  **String**.


## Syntax

_expression_. `FormulaLocal`

_expression_ A variable that represents a [Series](Excel.Series-graph-object.md) object.


## Remarks

If the cell contains a constant, this property returns that constant. If the cell is empty, the property returns an empty string. If the cell contains a formula, the property returns the formula as a string, in the same format in which it would be displayed in the formula bar (including the equal sign).

If you set the value or formula of a cell to a date, Microsoft Excel checks to see whether that cell is already formatted with one of the date or time number formats. If not, the number format is changed to the default short date number format.

If the range is a one- or two-dimensional range, you can set the formula to a Visual Basic array of the same dimensions. Similarly, you can put the formula into a Visual Basic array.

Setting the formula of a multiple-cell range fills all cells in the range with the formula.


## See also


[Series Object](Excel.Series(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]