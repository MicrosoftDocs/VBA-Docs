---
title: CalculatedMember object (Excel)
keywords: vbaxl10.chm685072
f1_keywords:
- vbaxl10.chm685072
ms.prod: excel
api_name:
- Excel.CalculatedMember
ms.assetid: 07a1f8df-107e-a5fd-3d15-dfc92916c4c6
ms.date: 03/29/2019
localization_priority: Normal
---


# CalculatedMember object (Excel)

Represents the calculated fields, calculated items, and named sets for PivotTables with Online Analytical Processing (OLAP) data sources.


## Remarks

Use the **Add** method or the **Item** property of the **[CalculatedMembers](Excel.CalculatedMembers.md)** collection to return a **CalculatedMember** object.

With a **CalculatedMember** object, you can check the validity of a calculated field or item in a PivotTable by using the **IsValid** property.

> [!NOTE] 
> The **IsValid** property returns **True** if the PivotTable is not currently connected to the data source. Use the **[MakeConnection](Excel.PivotCache.MakeConnection.md)** method of the **PivotCache** object before testing the **IsValid** property.


## Example

The following example notifies the user whether the calculated member is valid. This example assumes that a PivotTable exists on the active worksheet that contains either a valid or invalid calculated member.

```vb
Sub CheckValidity() 
 
 Dim pvtTable As PivotTable 
 Dim pvtCache As PivotCache 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtCache = Application.ActiveWorkbook.PivotCaches.Item(1) 
 
 ' Handle run-time error if external source is not an OLEDB data source. 
 On Error GoTo Not_OLEDB 
 
 ' Check connection setting and make connection if necessary. 
 If pvtCache.IsConnected = False Then 
 pvtCache.MakeConnection 
 End If 
 
 ' Check if calculated member is valid. 
 If pvtTable.CalculatedMembers.Item(1).IsValid = True Then 
 MsgBox "The calculated member is valid." 
 Else 
 MsgBox "The calculated member is not valid." 
 End If 
 
End Sub
```

## Methods

- [Delete](Excel.CalculatedMember.Delete.md)

## Properties

- [Application](Excel.CalculatedMember.Application.md)
- [Creator](Excel.CalculatedMember.Creator.md)
- [DisplayFolder](Excel.CalculatedMember.DisplayFolder.md)
- [Dynamic](Excel.CalculatedMember.Dynamic.md)
- [FlattenHierarchies](Excel.CalculatedMember.FlattenHierarchies.md)
- [Formula](Excel.CalculatedMember.Formula.md)
- [HierarchizeDistinct](Excel.CalculatedMember.HierarchizeDistinct.md)
- [IsValid](Excel.CalculatedMember.IsValid.md)
- [MeasureGroup](Excel.calculatedmember.measuregroup.md)
- [Name](Excel.CalculatedMember.Name.md)
- [NumberFormat](Excel.calculatedmember.numberformat.md)
- [Parent](Excel.CalculatedMember.Parent.md)
- [ParentHierarchy](Excel.calculatedmember.parenthierarchy.md)
- [ParentMember](Excel.calculatedmember.parentmember.md)
- [SolveOrder](Excel.CalculatedMember.SolveOrder.md)
- [SourceName](Excel.CalculatedMember.SourceName.md)
- [Type](Excel.CalculatedMember.Type.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]