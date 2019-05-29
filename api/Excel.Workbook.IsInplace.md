---
title: Workbook.IsInplace property (Excel)
keywords: vbaxl10.chm199184
f1_keywords:
- vbaxl10.chm199184
ms.prod: excel
api_name:
- Excel.Workbook.IsInplace
ms.assetid: f492c09f-79d1-cde0-6cf1-db9644e41589
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.IsInplace property (Excel)

**True** if the specified workbook is being edited in place. **False** if the workbook has been opened in Microsoft Excel for editing. Read-only **Boolean**.


## Syntax

_expression_.**IsInplace**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

This example indicates whether the workbook was opened for editing in place or in Microsoft Excel.

```vb
Private Sub Workbook_Open() 
 If ThisWorkbook.IsInPlace = True Then 
 MsgBox "Editing in place" 
 Else 
 MsgBox "Editing in Microsoft Excel" 
 End If 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]