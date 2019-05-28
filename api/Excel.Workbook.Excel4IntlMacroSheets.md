---
title: Workbook.Excel4IntlMacroSheets property (Excel)
keywords: vbaxl10.chm199169
f1_keywords:
- vbaxl10.chm199169
ms.prod: excel
api_name:
- Excel.Workbook.Excel4IntlMacroSheets
ms.assetid: 70a8c8d0-1169-7c3d-904e-5a32a4693f45
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.Excel4IntlMacroSheets property (Excel)

Returns a **[Sheets](Excel.Sheets.md)** collection that represents all the Microsoft Excel 4.0 international macro sheets in the specified workbook. Read-only.


## Syntax

_expression_.**Excel4IntlMacroSheets**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

This example displays the number of Microsoft Excel 4.0 international macro sheets in the active workbook.

```vb
MsgBox "There are " & _ 
 ActiveWorkbook.Excel4IntlMacroSheets.Count & _ 
 " Microsoft Excel 4.0 international macro sheets" & _ 
 " in this workbook."
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]