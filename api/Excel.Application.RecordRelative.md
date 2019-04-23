---
title: Application.RecordRelative property (Excel)
keywords: vbaxl10.chm133196
f1_keywords:
- vbaxl10.chm133196
ms.prod: excel
api_name:
- Excel.Application.RecordRelative
ms.assetid: 64e634e4-30e2-0794-1120-0960e32fe821
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.RecordRelative property (Excel)

**True** if macros are recorded by using relative references; **False** if recording is absolute. Read-only **Boolean**.


## Syntax

_expression_.**RecordRelative**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example displays the address of the active cell on Sheet1 in A1 style if **RecordRelative** is **False**; otherwise, it displays the address in R1C1 style.

```vb
Worksheets("Sheet1").Activate 
If Application.RecordRelative = False Then 
 MsgBox ActiveCell.Address(ReferenceStyle:=xlA1) 
Else 
 MsgBox ActiveCell.Address(ReferenceStyle:=xlR1C1) 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]