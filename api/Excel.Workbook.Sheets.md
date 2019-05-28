---
title: Workbook.Sheets property (Excel)
keywords: vbaxl10.chm199152
f1_keywords:
- vbaxl10.chm199152
ms.prod: excel
api_name:
- Excel.Workbook.Sheets
ms.assetid: 45e4e19e-55ea-9615-231d-9435ba6d5a63
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.Sheets property (Excel)

Returns a **[Sheets](Excel.Sheets.md)** collection that represents all the sheets in the specified workbook. Read-only **Sheets** object.


## Syntax

_expression_.**Sheets**

_expression_ An expression that returns a **[Workbook](Excel.Workbook.md)** object.


## Remarks

Using this property without an object qualifier is equivalent to using **ActiveWorkbook.Sheets**.


## Example

This example creates a new worksheet and then places a list of the active workbook's sheet names in the first column.

```vb
Set newSheet = Sheets.Add(Type:=xlWorksheet) 
For i = 1 To Sheets.Count 
 newSheet.Cells(i, 1).Value = Sheets(i).Name 
Next i
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
