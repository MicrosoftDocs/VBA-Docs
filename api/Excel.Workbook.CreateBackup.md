---
title: Workbook.CreateBackup property (Excel)
keywords: vbaxl10.chm199093
f1_keywords:
- vbaxl10.chm199093
ms.prod: excel
api_name:
- Excel.Workbook.CreateBackup
ms.assetid: 33f05bf8-00ef-81f4-c083-30326f019cd4
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.CreateBackup property (Excel)

**True** if a backup file is created when this file is saved. Read-only **Boolean**.


## Syntax

_expression_.**CreateBackup**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

This example displays a message if a backup file is created when the active workbook is saved.

```vb
If ActiveWorkbook.CreateBackup = True Then 
 MsgBox "Remember, there is a backup copy of this workbook" 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]