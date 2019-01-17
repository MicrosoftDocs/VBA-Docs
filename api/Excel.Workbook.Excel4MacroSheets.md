---
title: Workbook.Excel4MacroSheets property (Excel)
keywords: vbaxl10.chm199170
f1_keywords:
- vbaxl10.chm199170
ms.prod: excel
api_name:
- Excel.Workbook.Excel4MacroSheets
ms.assetid: 29161ab8-da75-c7b5-561d-f4423b8ab1ef
ms.date: 06/08/2017
localization_priority: Normal
---


# Workbook.Excel4MacroSheets property (Excel)

Returns a  **[Sheets](Excel.Sheets.md)** collection that represents all the Microsoft Excel 4.0 macro sheets in the specified workbook. Read-only.


## Syntax

_expression_. `Excel4MacroSheets`

_expression_ A variable that represents a [Workbook](./Excel.Workbook.md) object.


## Remarks

Using this property with the  **Application** object or without an object qualifier is equivalent to using `ActiveWorkbook.Excel4MacroSheets`.


## Example

This example displays the number of Microsoft Excel 4.0 macro sheets in the active workbook.


```vb
MsgBox "There are " & ActiveWorkbook.Excel4MacroSheets.Count & _ 
 " Microsoft Excel 4.0 macro sheets in this workbook."
```


## See also


[Workbook Object](Excel.Workbook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]