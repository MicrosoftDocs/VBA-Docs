---
title: CalculatedMembers Object (Excel)
keywords: vbaxl10.chm683072
f1_keywords:
- vbaxl10.chm683072
ms.prod: excel
api_name:
- Excel.CalculatedMembers
ms.assetid: 3c664ac6-e2f8-f631-006d-6a16c380641e
ms.date: 06/08/2017
---


# CalculatedMembers Object (Excel)

A collection of all the  **[CalculatedMember](Excel.CalculatedMembers.md)** objects on the specified PivotTable.


## Remarks

 Each **CalculatedMember** object represents a calculated member or calculated measure.

Use the  **[CalculatedMembers](Excel.PivotTable.CalculatedMembers.md)** property of the **[PivotTable](Excel.PivotTable.md)** object to return a **CalculatedMembers** collection.

There are three supported types of calculated members:  _Named Sets_ , _Calculated Measures_ , and _Calculated Members_ . Object model support has been available for all three of these types since Excel 2010. User interface support was made available for Named Sets in Excel 2010. In Excel 2013, the OLAP Calculated Members and Calculated Measures feature was created to build a user interface for the calculated members and measures object model.

 **Named Sets** are used exactly the same as in Excel 2010. Named Sets should continue to use the method CalculatedMembers.[CalculatedMembers.Add Method (Excel)](Excel.CalculatedMembers.Add.md) and the type[XlCalculatedMemberType Enumeration (Excel)](Excel.XlCalculatedMemberType.md).

 **Calculated Members** have the following changes for Excel 2013:


- They now use the method called CalculatedMembers.[CalculatedMembers.AddCalculatedMember Method (Excel)](Excel.calculatedmembers.addcalculatedmember.md).
    
- They support the property [CalculatedMember.ParentHierarchy Property (Excel)](Excel.calculatedmember.parenthierarchy.md).
    
- They support the property [CalculatedMember.ParentMember Property (Excel)](Excel.calculatedmember.parentmember.md).
    
- They support the property [CalculatedMember.NumberFormat Property (Excel)](Excel.calculatedmember.numberformat.md).
    
 **Calculated Measures** have the following changes for Excel 2013:


- They now use the method called CalculatedMembers.[CalculatedMembers.AddCalculatedMember Method (Excel)](Excel.calculatedmembers.addcalculatedmember.md).
    
- They now use the type [XlCalculatedMemberType Enumeration (Excel)](Excel.XlCalculatedMemberType.md).
    
- They support the property [CalculatedMember.DisplayFolder Property (Excel)](Excel.CalculatedMember.DisplayFolder.md).
    
- They support the property [CalculatedMember.NumberFormat Property (Excel)](Excel.calculatedmember.numberformat.md).
    

## Example

The following example adds a set to a PivotTable, assuming a PivotTable from the FoodMart SQL database exists on the active worksheet.


```vb
Sub UseCalculatedMember() 
 Dim pvtTable As PivotTable 
 Set pvtTable = ActiveSheet.PivotTables(1)
 pvtTable.CalculatedMembers.Add Name:="[Beef]", _ 
 Formula:="'{[Product].[All Products].Children}'", _ 
 Type:=xlCalculatedSet 
 
End Sub
```


 **Note**  For the  **Add** method in the previous example, the **Formula** argument must have a valid MDX syntax statement. The **Name** argument has to be acceptable to the Online Analytical Processing (OLAP) provider and the **Type** argument has to be defined.


## See also


[Excel Object Model Reference](./overview/object-model-excel-vba-reference.md)


