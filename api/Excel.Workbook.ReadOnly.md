---
title: Workbook.ReadOnly property (Excel)
keywords: vbaxl10.chm199133
f1_keywords:
- vbaxl10.chm199133
ms.prod: excel
api_name:
- Excel.Workbook.ReadOnly
ms.assetid: f3c0ec74-63af-ed76-f854-ce2382b9fcf3
ms.date: 06/08/2017
localization_priority: Priority
---


# Workbook.ReadOnly property (Excel)

 Returns **True** if the object has been opened as read-only. Read-only **Boolean**.


## Syntax

_expression_. `ReadOnly`

_expression_ A variable that represents a [Workbook](./Excel.Workbook.md) object.


## Example

If the active workbook is read-only, this example saves it as Newfile.xls.


```vb
If ActiveWorkbook.ReadOnly Then 
 ActiveWorkbook.SaveAs fileName:="NEWFILE.XLS" 
End If
```


## See also


[Workbook Object](Excel.Workbook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]