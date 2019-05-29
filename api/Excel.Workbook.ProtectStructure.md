---
title: Workbook.ProtectStructure property (Excel)
keywords: vbaxl10.chm199131
f1_keywords:
- vbaxl10.chm199131
ms.prod: excel
api_name:
- Excel.Workbook.ProtectStructure
ms.assetid: bf721b60-0ad1-f71c-7ef4-74d2196d320e
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.ProtectStructure property (Excel)

**True** if the order of the sheets in the workbook is protected. Read-only **Boolean**.


## Syntax

_expression_.**ProtectStructure**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

This example displays a message if the order of the sheets in the active workbook is protected.

```vb
If ActiveWorkbook.ProtectStructure = True Then 
 MsgBox "Remember, you cannot delete, add, or change " & _ 
 Chr(13) & _ 
 "the location of any sheets in this workbook." 
End If
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]