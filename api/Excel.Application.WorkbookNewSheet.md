---
title: Application.WorkbookNewSheet event (Excel)
keywords: vbaxl10.chm504087
f1_keywords:
- vbaxl10.chm504087
ms.prod: excel
api_name:
- Excel.Application.WorkbookNewSheet
ms.assetid: 5190254f-b7f4-10e5-41f5-704b1466ff68
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.WorkbookNewSheet event (Excel)

Occurs when a new sheet is created in any open workbook.


## Syntax

_expression_.**WorkbookNewSheet** (_Wb_, _Sh_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The workbook.|
| _Sh_|Required| **Object**|The new sheet.|

## Return value

Nothing


## Example

This example moves the new sheet to the end of the workbook.

```vb
Private Sub App_WorkbookNewSheet(ByVal Wb As Workbook, _ 
 ByVal Sh As Object) 
 Sh.Move After:=Wb.Sheets(Wb.Sheets.Count) 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]