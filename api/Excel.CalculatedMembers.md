---
title: CalculatedMembers object (Excel)
keywords: vbaxl10.chm683072
f1_keywords:
- vbaxl10.chm683072
ms.prod: excel
api_name:
- Excel.CalculatedMembers
ms.assetid: 3c664ac6-e2f8-f631-006d-6a16c380641e
ms.date: 03/29/2019
localization_priority: Normal
---


# CalculatedMembers object (Excel)

A collection of all the **[CalculatedMember](Excel.CalculatedMembers.md)** objects on the specified PivotTable.


## Remarks

Each **CalculatedMember** object represents a calculated member or calculated measure.

Use the **[CalculatedMembers](Excel.PivotTable.CalculatedMembers.md)** property of the **PivotTable** object to return a **CalculatedMembers** collection.

There are three supported types of calculated members: **Named Sets**, **Calculated Members**, and **Calculated Measures**. Object model support has been available for all three types since Excel 2010. User interface support was made available for **Named Sets** in Excel 2010. In Excel 2013, the OLAP **Calculated Members and Calculated Measures** feature was created to build a user interface for the calculated members and measures object model.

**Named Sets** is used exactly the same as in Excel 2010. **Named Sets** should continue to use the **Add** method, and the type **[XlCalculatedMemberType](excel.xlcalculatedmembertype.md)** enumeration.

**Calculated Members** has the following changes for Excel 2013:

- It now uses the **AddCalculatedMember** method.
    
- It supports the following properties of the **CalculatedMember** object:

  - **[ParentHierarchy](Excel.calculatedmember.parenthierarchy.md)** property
    
  - **[ParentMember](Excel.calculatedmember.parentmember.md)** property 
    
  - **[NumberFormat](Excel.calculatedmember.numberformat.md)** property 
    
**Calculated Measures** has the following changes for Excel 2013:

- It now uses the **AddCalculatedMember** method.
    
- It now uses the type **[XlCalculatedMemberType](Excel.xlCalculatedMemberType.md)** enumeration.

- It supports the following properties of the **CalculatedMember** object:
    
  - **[DisplayFolder](Excel.CalculatedMember.DisplayFolder.md)** property
    
  - **[NumberFormat](Excel.calculatedmember.numberformat.md)** property 
    

## Example

The following example adds a set to a PivotTable, assuming that a PivotTable from the FoodMart SQL database exists on the active worksheet.

```vb
Sub UseCalculatedMember() 
 Dim pvtTable As PivotTable 
 Set pvtTable = ActiveSheet.PivotTables(1)
 pvtTable.CalculatedMembers.Add Name:="[Beef]", _ 
 Formula:="'{[Product].[All Products].Children}'", _ 
 Type:=xlCalculatedSet 
 
End Sub
```

> [!NOTE] 
> For the **Add** method in the previous example, the **Formula** argument must have a valid MDX syntax statement. The **Name** argument has to be acceptable to the Online Analytical Processing (OLAP) provider and the **Type** argument has to be defined.


## Methods

- [Add](Excel.CalculatedMembers.Add.md)
- [AddCalculatedMember](Excel.calculatedmembers.addcalculatedmember.md)

## Properties

- [Application](Excel.CalculatedMembers.Application.md)
- [Count](Excel.CalculatedMembers.Count.md)
- [Creator](Excel.CalculatedMembers.Creator.md)
- [Item](Excel.CalculatedMembers.Item.md)
- [Parent](Excel.CalculatedMembers.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]