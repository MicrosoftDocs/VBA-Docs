---
title: Worksheet.Activate event (Excel)
keywords: vbaxl10.chm502076
f1_keywords:
- vbaxl10.chm502076
ms.prod: excel
api_name:
- Excel.Worksheet.Activate
ms.assetid: 4fac262c-ea1a-1d2f-bd02-0537c843198c
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.Activate event (Excel)

Occurs when a workbook, worksheet, chart sheet, or embedded chart is activated.


## Syntax

_expression_.**Activate**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Return value

**Nothing**


## Remarks

This event doesn't occur when you create a new window.

When you switch between two windows showing the same workbook, the **[WindowActivate](excel.workbook.windowactivate.md)** event occurs, but the **Activate** event for the workbook doesn't occur.


## Example

This example sorts the range A1:A10 when the worksheet is activated.

```vb
Private Sub Worksheet_Activate() 
 Me.Range("a1:a10").Sort Key1:=Range("a1"), Order1:=xlAscending 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
