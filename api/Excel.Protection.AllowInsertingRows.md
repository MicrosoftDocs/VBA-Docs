---
title: Protection.AllowInsertingRows property (Excel)
keywords: vbaxl10.chm719077
f1_keywords:
- vbaxl10.chm719077
ms.prod: excel
api_name:
- Excel.Protection.AllowInsertingRows
ms.assetid: 481fb5d0-31c9-9c28-c5a0-3f3abc48ad3a
ms.date: 05/09/2019
localization_priority: Normal
---


# Protection.AllowInsertingRows property (Excel)

Returns **True** if the insertion of rows is allowed on a protected worksheet. Read-only **Boolean**.


## Syntax

_expression_.**AllowInsertingRows**

_expression_ A variable that represents a **[Protection](Excel.Protection.md)** object.


## Remarks

The **AllowInsertingRows** property can be set by using the **[Protect](Excel.Worksheet.Protect.md)** method arguments.


## Example

This example allows the user to insert rows on the protected worksheet and notifies the user.

```vb
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 ' Allow rows to be inserted on a protected worksheet. 
 If ActiveSheet.Protection.AllowInsertingRows = False Then 
 ActiveSheet.Protect AllowInsertingRows:=True 
 End If 
 
 MsgBox "Rows can be inserted on this protected worksheet." 
 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]