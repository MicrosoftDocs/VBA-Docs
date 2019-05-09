---
title: Protection.AllowUsingPivotTables property (Excel)
keywords: vbaxl10.chm719083
f1_keywords:
- vbaxl10.chm719083
ms.prod: excel
api_name:
- Excel.Protection.AllowUsingPivotTables
ms.assetid: 42968839-1d82-3c0e-172b-1389c772f9a1
ms.date: 05/09/2019
localization_priority: Normal
---


# Protection.AllowUsingPivotTables property (Excel)

Returns **True** if the user is allowed to manipulate PivotTables on a protected worksheet. Read-only **Boolean**.


## Syntax

_expression_.**AllowUsingPivotTables**

_expression_ A variable that represents a **[Protection](Excel.Protection.md)** object.


## Remarks

The **AllowUsingPivotTables** property applies to non-OLAP source data.

The **AllowUsingPivotTables** property can be set by using the **[Protect](Excel.Worksheet.Protect.md)** method arguments.


## Example

This example allows the user to access the PivotTable report and notifies the user. It assumes that a non-OLAP PivotTable report exists on the active worksheet.

```vb
Sub ProtectionOptions() 
 
 ActiveSheet.Unprotect 
 
 ' Allow PivotTables to be manipulated on a protected worksheet. 
 If ActiveSheet.Protection.Allow UsingPivotTables = False Then 
 ActiveSheet.Protect AllowUsingPivotTables:=True 
 End If 
 
 MsgBox "PivotTables can be manipulated on the protected worksheet." 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]