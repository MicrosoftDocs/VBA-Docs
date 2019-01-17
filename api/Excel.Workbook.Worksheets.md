---
title: Workbook.Worksheets property (Excel)
keywords: vbaxl10.chm199166
f1_keywords:
- vbaxl10.chm199166
ms.prod: excel
api_name:
- Excel.Workbook.Worksheets
ms.assetid: 8b7d660d-ca49-0bd0-dc57-64defa47bd5e
ms.date: 06/08/2017
localization_priority: Priority
---


# Workbook.Worksheets property (Excel)

Returns a  **[Sheets](Excel.Sheets.md)** collection that represents all the worksheets in the specified workbook. Read-only **Sheets** object.


## Syntax

_expression_. `Worksheets`

_expression_ A variable that represents a [Workbook](./Excel.Workbook.md) object.


## Remarks

Using this property without an object qualifier returns all the worksheets in the active workbook.

This property doesn't return macro sheets; use the  **[Excel4MacroSheets](Excel.Workbook.Excel4MacroSheets.md)** property or the **[Excel4IntlMacroSheets](Excel.Workbook.Excel4IntlMacroSheets.md)** property to return those sheets.


## Example

This example displays the value in cell A1 on Sheet1 in the active workbook.


```vb
MsgBox Worksheets("Sheet1").Range("A1").Value
```

This example displays the name of each worksheet in the active workbook.




```vb
For Each ws In Worksheets 
 MsgBox ws.Name 
Next ws
```

This example adds a new worksheet to the active workbook and then sets the name of the worksheet.




```vb
Set newSheet = Worksheets.Add 
newSheet.Name = "current Budget"
```


## See also


[Workbook Object](Excel.Workbook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]