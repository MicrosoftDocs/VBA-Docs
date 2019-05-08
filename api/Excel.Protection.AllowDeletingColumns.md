---
title: Protection.AllowDeletingColumns property (Excel)
keywords: vbaxl10.chm719079
f1_keywords:
- vbaxl10.chm719079
ms.prod: excel
api_name:
- Excel.Protection.AllowDeletingColumns
ms.assetid: 602e0599-f444-0e81-9d9c-70f1f8093a29
ms.date: 05/09/2019
localization_priority: Normal
---


# Protection.AllowDeletingColumns property (Excel)

Returns **True** if the deletion of columns is allowed on a protected worksheet. Read-only **Boolean**.


## Syntax

_expression_.**AllowDeletingColumns**

_expression_ A variable that represents a **[Protection](Excel.Protection.md)** object.


## Remarks

The **AllowDeletingColumns** property can be set by using the **[Protect](Excel.Worksheet.Protect.md)** method arguments.

The columns containing the cells to be deleted must be unlocked when the sheet is protected.


## Example

This example unlocks column A, and then allows the user to delete column A on the protected worksheet.

```vb
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 'Unlock column A. 
 Columns("A:A").Locked = False 
 
 ' Allow column A to be deleted on a protected worksheet. 
 If ActiveSheet.Protection.AllowDeletingColumns = False Then 
 ActiveSheet.Protect AllowDeletingColumns:=True 
 End If 
 
 MsgBox "Column A can be deleted on this protected worksheet." 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]