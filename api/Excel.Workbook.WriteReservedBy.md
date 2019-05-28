---
title: Workbook.WriteReservedBy property (Excel)
keywords: vbaxl10.chm199168
f1_keywords:
- vbaxl10.chm199168
ms.prod: excel
api_name:
- Excel.Workbook.WriteReservedBy
ms.assetid: f053c197-3af3-9ab7-bee1-f72ee311a5b8
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.WriteReservedBy property (Excel)

Returns the name of the user who currently has write permission for the workbook. Read-only **String**.


## Syntax

_expression_.**WriteReservedBy**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

If the active workbook is write-reserved, this example displays a message that contains the name of the user who saved the workbook as write-reserved.

```vb
With ActiveWorkbook 
 If .WriteReserved = True Then 
 MsgBox "Please contact " & .WriteReservedBy & Chr(13) & _ 
 " if you need to insert data in this workbook." 
 End If 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]