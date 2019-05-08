---
title: Protection.AllowFormattingColumns property (Excel)
keywords: vbaxl10.chm719074
f1_keywords:
- vbaxl10.chm719074
ms.prod: excel
api_name:
- Excel.Protection.AllowFormattingColumns
ms.assetid: 1cdfeea0-5c5e-1f6c-47c7-a351bb6745b7
ms.date: 05/09/2019
localization_priority: Normal
---


# Protection.AllowFormattingColumns property (Excel)

Returns **True** if the formatting of columns is allowed on a protected worksheet. Read-only **Boolean**.


## Syntax

_expression_.**AllowFormattingColumns**

_expression_ A variable that represents a **[Protection](Excel.Protection.md)** object.


## Remarks

The **AllowFormattingColumns** property can be set by using the **[Protect](Excel.Worksheet.Protect.md)** method arguments.


## Example

This example allows the user to format columns on the protected worksheet and notifies the user.

```vb
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 ' Allow columns to be formatted on a protected worksheet. 
 If ActiveSheet.Protection.AllowFormattingColumns = False Then 
 ActiveSheet.Protect AllowFormattingColumns:=True 
 End If 
 
 MsgBox "Columns can be formatted on this protected worksheet." 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]