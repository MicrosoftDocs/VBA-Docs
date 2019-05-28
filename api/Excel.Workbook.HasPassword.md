---
title: Workbook.HasPassword property (Excel)
keywords: vbaxl10.chm199104
f1_keywords:
- vbaxl10.chm199104
ms.prod: excel
api_name:
- Excel.Workbook.HasPassword
ms.assetid: e3cfdc90-1e82-5556-0064-e8269ba92539
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.HasPassword property (Excel)

**True** if the workbook has a protection password. Read-only **Boolean**.


## Syntax

_expression_.**HasPassword**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Remarks

You can assign a protection password to a workbook by using the **[SaveAs](Excel.Workbook.SaveAs.md)** method.


## Example

This example displays a message if the active workbook has a protection password.

```vb
If ActiveWorkbook.HasPassword = True Then 
 MsgBox "Remember to obtain the workbook password" & Chr(13) & _ 
 " from the Network Administrator." 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]