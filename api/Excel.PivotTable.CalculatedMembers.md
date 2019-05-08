---
title: PivotTable.CalculatedMembers property (Excel)
keywords: vbaxl10.chm235145
f1_keywords:
- vbaxl10.chm235145
ms.prod: excel
api_name:
- Excel.PivotTable.CalculatedMembers
ms.assetid: 65e7ffd6-e01d-f8fc-3adb-a1bcb1046fcf
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.CalculatedMembers property (Excel)

Returns a **[CalculatedMembers](Excel.CalculatedMembers.md)** collection representing all the calculated members and calculated measures for an OLAP PivotTable.


## Syntax

_expression_.**CalculatedMembers**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

This property is used for Online Analytical Processing (OLAP) sources; a non-OLAP source will return a run-time error.


## Example

This example adds a set to the PivotTable. It assumes that a PivotTable exists on the active worksheet that is connected to an OLAP data source that contains a field titled [Product].[All Products].

```vb
Sub UseCalculatedMember() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Add the calculated member. 
 pvtTable.CalculatedMembers.Add Name:="[Beef]", _ 
 Formula:="'{[Product].[All Products].Children}'", _ 
 Type:=xlCalculatedSet 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]