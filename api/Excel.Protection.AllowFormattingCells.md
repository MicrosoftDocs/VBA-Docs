---
title: Protection.AllowFormattingCells property (Excel)
keywords: vbaxl10.chm719073
f1_keywords:
- vbaxl10.chm719073
ms.prod: excel
api_name:
- Excel.Protection.AllowFormattingCells
ms.assetid: 6e3d6fd1-a1f5-95c1-0ef2-795eba31b904
ms.date: 05/09/2019
localization_priority: Normal
---


# Protection.AllowFormattingCells property (Excel)

Returns **True** if the formatting of cells is allowed on a protected worksheet. Read-only **Boolean**.


## Syntax

_expression_.**AllowFormattingCells**

_expression_ A variable that represents a **[Protection](Excel.Protection.md)** object.


## Remarks

The **AllowFormattingCells** property can be set by using the **[Protect](Excel.Worksheet.Protect.md)** method arguments.

Use of this property disables the protection tab, allowing the user to change all formats, but not to unlock or unhide ranges.


## Example

This example allows the user to format cells on the protected worksheet and notifies the user.

```vb
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 ' Allow cells to be formatted on a protected worksheet. 
 If ActiveSheet.Protection.AllowFormattingCells = False Then 
 ActiveSheet.Protect AllowFormattingCells:=True 
 End If 
 
 MsgBox "Cells can be formatted on this protected worksheet." 
 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
