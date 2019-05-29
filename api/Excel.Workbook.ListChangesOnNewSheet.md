---
title: Workbook.ListChangesOnNewSheet property (Excel)
keywords: vbaxl10.chm199175
f1_keywords:
- vbaxl10.chm199175
ms.prod: excel
api_name:
- Excel.Workbook.ListChangesOnNewSheet
ms.assetid: 77adf429-baa5-f2be-6139-c2b07dda5174
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.ListChangesOnNewSheet property (Excel)

**True** if changes to the shared workbook are shown on a separate worksheet. Read/write **Boolean**.


## Syntax

_expression_.**ListChangesOnNewSheet**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

This example shows changes to the shared workbook on a separate worksheet.

```vb
With ActiveWorkbook 
 .HighlightChangesOptions _ 
 When:=xlSinceMyLastSave, _ 
 Who:="Everyone" 
 .ListChangesOnNewSheet = True 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]