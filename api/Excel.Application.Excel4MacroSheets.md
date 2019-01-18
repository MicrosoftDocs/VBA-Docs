---
title: Application.Excel4MacroSheets property (Excel)
keywords: vbaxl10.chm132118
f1_keywords:
- vbaxl10.chm132118
ms.prod: excel
api_name:
- Excel.Application.Excel4MacroSheets
ms.assetid: d1ee907a-302c-4bd5-5455-75c328f94268
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Excel4MacroSheets property (Excel)

Returns a  **[Sheets](Excel.Sheets.md)** collection that represents all the Microsoft Excel 4.0 macro sheets in the specified workbook. Read-only.


## Syntax

_expression_. `Excel4MacroSheets`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Remarks

Using this property with the  **Application** object or without an object qualifier is equivalent to using `ActiveWorkbook.Excel4MacroSheets`.


## Example

This example displays the number of Microsoft Excel 4.0 macro sheets in the active workbook.


```vb
MsgBox "There are " & ActiveWorkbook.Excel4MacroSheets.Count & _ 
 " Microsoft Excel 4.0 macro sheets in this workbook."
```


## See also


[Application Object](Excel.Application(object).md)

