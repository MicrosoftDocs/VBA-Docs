---
title: PivotItem object (Excel)
keywords: vbaxl10.chm245072
f1_keywords:
- vbaxl10.chm245072
ms.prod: excel
api_name:
- Excel.PivotItem
ms.assetid: 5829a1d9-0924-9ce8-1120-229e4595285a
ms.date: 03/30/2019
localization_priority: Normal
---


# PivotItem object (Excel)

Represents an item in a PivotTable field.


## Remarks

The items are the individual data entries in a field category. The **PivotItem** object is a member of the **[PivotItems](Excel.PivotItems.md)** collection. The **PivotItems** collection contains all the items in a **PivotField** object.


## Example

Use **[PivotItems](Excel.PivotField.PivotItems.md)** (_index_), where _index_ is the item index number or name, to return a single **PivotItem** object. 

The following example hides all entries in the first PivotTable report on Sheet3 that contain "1998" in the Year field.

```vb
Worksheets("sheet3").PivotTables(1) _ 
 .PivotFields("year").PivotItems("1998").Visible = False
```


## Methods

- [Delete](Excel.PivotItem.Delete.md)
- [DrillTo](Excel.PivotItem.DrillTo.md)

## Properties

- [Application](Excel.PivotItem.Application.md)
- [Caption](Excel.PivotItem.Caption.md)
- [ChildItems](Excel.PivotItem.ChildItems.md)
- [Creator](Excel.PivotItem.Creator.md)
- [DataRange](Excel.PivotItem.DataRange.md)
- [DrilledDown](Excel.PivotItem.DrilledDown.md)
- [Formula](Excel.PivotItem.Formula.md)
- [IsCalculated](Excel.PivotItem.IsCalculated.md)
- [LabelRange](Excel.PivotItem.LabelRange.md)
- [Name](Excel.PivotItem.Name.md)
- [Parent](Excel.PivotItem.Parent.md)
- [ParentItem](Excel.PivotItem.ParentItem.md)
- [ParentShowDetail](Excel.PivotItem.ParentShowDetail.md)
- [Position](Excel.PivotItem.Position.md)
- [RecordCount](Excel.PivotItem.RecordCount.md)
- [ShowDetail](Excel.PivotItem.ShowDetail.md)
- [SourceName](Excel.PivotItem.SourceName.md)
- [SourceNameStandard](Excel.PivotItem.SourceNameStandard.md)
- [StandardFormula](Excel.PivotItem.StandardFormula.md)
- [Value](Excel.PivotItem.Value.md)
- [Visible](Excel.PivotItem.Visible.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]