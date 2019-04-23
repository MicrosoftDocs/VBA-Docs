---
title: CalculatedMember.SolveOrder property (Excel)
keywords: vbaxl10.chm686076
f1_keywords:
- vbaxl10.chm686076
ms.prod: excel
api_name:
- Excel.CalculatedMember.SolveOrder
ms.assetid: 45e461ac-4900-000b-cb72-4022bcc1a7c9
ms.date: 04/13/2019
localization_priority: Normal
---


# CalculatedMember.SolveOrder property (Excel)

Returns a **Long** specifying the value of the calculated member's solve order MDX (Mulitdimensional Expression) argument. The default value is zero. Read-only.


## Syntax

_expression_.**SolveOrder**

_expression_ A variable that represents a **[CalculatedMember](Excel.CalculatedMember.md)** object.


## Example

This example determines the solve order for a calculated member and notifies the user. The example assumes that a PivotTable exists on the active worksheet.

```vb
Sub CheckSolveOrder() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Determine solve order and notify user. 
 If pvtTable.CalculatedMembers.Item(1).SolveOrder = 0 Then 
 MsgBox "The solve order is set to the default value." 
 Else 
 MsgBox "The solve order is not set to the default value." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]