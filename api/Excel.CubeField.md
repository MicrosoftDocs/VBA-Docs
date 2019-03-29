---
title: CubeField object (Excel)
keywords: vbaxl10.chm667072
f1_keywords:
- vbaxl10.chm667072
ms.prod: excel
api_name:
- Excel.CubeField
ms.assetid: 6db16910-6c27-651a-c388-e54e27fe4519
ms.date: 03/29/2019
localization_priority: Normal
---


# CubeField object (Excel)

Represents a hierarchy or measure field from an OLAP cube. In a PivotTable report, the **CubeField** object is a member of the **[CubeFields](Excel.CubeFields.md)** collection.


## Example

Use the **[CubeField](Excel.PivotField.CubeField.md)** property of the **PivotField** object to return the **CubeField** object. This example creates a list of the cube field names for all the hierarchy fields in the first OLAP-based PivotTable report on **Sheet1**.

```vb
Set objNewSheet = Worksheets.Add 
objNewSheet.Activate 
intRow = 1 
For Each objPF in _ 
 Worksheets("Sheet1").PivotTables(1).PivotFields 
 If objPF.CubeField.CubeFieldType = xlHierarchy Then 
 objNewSheet.Cells(intRow, 1).Value = objPF.Name 
 intRow = intRow + 1 
 End If 
Next objPF
```

<br/>

Use **CubeFields** (_index_), where _index_ is the cube field's index number, to return a single **CubeField** object. The following example determines the name of the second cube field in the first PivotTable report on the active worksheet.

```vb
strAlphaName = _ 
 ActiveSheet.PivotTables(1).CubeFields(2).Name
```


## Methods

- [AddMemberPropertyField](Excel.CubeField.AddMemberPropertyField.md)
- [AutoGroup](Excel.cubefield.autogroup.md)
- [ClearManualFilter](Excel.CubeField.ClearManualFilter.md)
- [CreatePivotFields](Excel.CubeField.CreatePivotFields.md)
- [Delete](Excel.CubeField.Delete.md)

## Properties

- [AllItemsVisible](Excel.CubeField.AllItemsVisible.md)
- [Application](Excel.CubeField.Application.md)
- [Caption](Excel.CubeField.Caption.md)
- [Creator](Excel.CubeField.Creator.md)
- [CubeFieldSubType](Excel.CubeField.CubeFieldSubType.md)
- [CubeFieldType](Excel.CubeField.CubeFieldType.md)
- [CurrentPageName](Excel.CubeField.CurrentPageName.md)
- [DragToColumn](Excel.CubeField.DragToColumn.md)
- [DragToData](Excel.CubeField.DragToData.md)
- [DragToHide](Excel.CubeField.DragToHide.md)
- [DragToPage](Excel.CubeField.DragToPage.md)
- [DragToRow](Excel.CubeField.DragToRow.md)
- [EnableMultiplePageItems](Excel.CubeField.EnableMultiplePageItems.md)
- [FlattenHierarchies](Excel.CubeField.FlattenHierarchies.md)
- [HasMemberProperties](Excel.CubeField.HasMemberProperties.md)
- [HierarchizeDistinct](Excel.CubeField.HierarchizeDistinct.md)
- [IncludeNewItemsInFilter](Excel.CubeField.IncludeNewItemsInFilter.md)
- [IsDate](Excel.CubeField.IsDate.md)
- [LayoutForm](Excel.CubeField.LayoutForm.md)
- [LayoutSubtotalLocation](Excel.CubeField.LayoutSubtotalLocation.md)
- [Name](Excel.CubeField.Name.md)
- [Orientation](Excel.CubeField.Orientation.md)
- [Parent](Excel.CubeField.Parent.md)
- [PivotFields](Excel.CubeField.PivotFields.md)
- [Position](Excel.CubeField.Position.md)
- [ShowInFieldList](Excel.CubeField.ShowInFieldList.md)
- [TreeviewControl](Excel.CubeField.TreeviewControl.md)
- [Value](Excel.CubeField.Value.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]