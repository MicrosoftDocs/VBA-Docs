---
title: Application.Cells property (Excel)
keywords: vbaxl10.chm183085
f1_keywords:
- vbaxl10.chm183085
ms.prod: excel
api_name:
- Excel.Application.Cells
ms.assetid: 9788c893-13c3-eb57-bcf7-50806b476ba3
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.Cells property (Excel)

Returns a **[Range](Excel.Range(object).md)** object that represents all the cells on the active worksheet. If the active document is not a worksheet, this property fails.


## Syntax

_expression_.**Cells**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

Because the **[Item](Excel.Range.Item.md)** property is the default property for the **Range** object, you can specify the row and column index immediately after the **Cells** keyword. For more information, see the **Item** property and the examples for this topic.

Using this property without an object qualifier returns a **Range** object that represents all the cells on the active worksheet.


## Example

This example looks at data in each row and inserts a blank row each time the value in column A changes.

```vb
Sub ChangeInsertRows()
    Application.ScreenUpdating = False
    Dim xRow As Long
    
    For xRow = Application.Cells(Rows.Count, 1).End(xlUp).Row To 3 Step -1
        If Application.Cells(xRow, 1).Value <> Application.Cells(xRow - 1, 1).Value Then Rows(xRow).Resize(1).Insert
    Next xRow
    
    Application.ScreenUpdating = True
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
