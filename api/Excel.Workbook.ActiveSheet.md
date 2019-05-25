---
title: Workbook.ActiveSheet property (Excel)
keywords: vbaxl10.chm199076
f1_keywords:
- vbaxl10.chm199076
ms.prod: excel
api_name:
- Excel.Workbook.ActiveSheet
ms.assetid: fb5578c3-64a7-edb7-4004-e608739d4c1e
ms.date: 05/25/2019
localization_priority: Normal
---


# Workbook.ActiveSheet property (Excel)

Returns a **[Worksheet](excel.worksheet.md)** object that represents the active sheet (the sheet on top) in the active workbook or specified workbook. Returns **Nothing** if no sheet is active.


## Syntax

_expression_.**ActiveSheet**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Remarks

Using the **ActiveSheet** property without an object qualifier returns the active sheet in the active workbook in the active window.

If a workbook appears in more than one window, the active sheet might be different in different windows.

## Example

This example displays the name of the active sheet.

```vb
MsgBox "The name of the active sheet is " & ActiveSheet.Name
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
