---
title: PivotItem.Formula property (Excel)
keywords: vbaxl10.chm246089
f1_keywords:
- vbaxl10.chm246089
ms.prod: excel
api_name:
- Excel.PivotItem.Formula
ms.assetid: c4e6a447-c910-79e5-701a-4f17210b7fb1
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotItem.Formula property (Excel)

Returns or sets a  **String** value that represents the object's formula in A1-style notation and in the language of the macro.


## Syntax

_expression_. `Formula`

_expression_ A variable that represents a [PivotItem](Excel.PivotItem.md) object.


## Remarks

This property is not available for OLAP data sources.

If the cell contains a constant, this property returns the constant. If the cell is empty, this property returns an empty string. If the cell contains a formula, the  **Formula** property returns the formula as a string in the same format that would be displayed in the formula bar (including the equal sign).

If you set the value or formula of a cell to a date, Microsoft Excel checks to see whether that cell is already formatted with one of the date or time number formats. If not, Microsoft Excel changes the number format to the default short date number format.

If the range is a one- or two-dimensional range, you can set the formula to a Visual Basic array of the same dimensions. Similarly, you can put the formula into a Visual Basic array.

Setting the formula for a multiple-cell range fills all cells in the range with the formula.


## See also


[PivotItem Object](Excel.PivotItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]