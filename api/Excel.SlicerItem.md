---
title: SlicerItem object (Excel)
keywords: vbaxl10.chm906072
f1_keywords:
- vbaxl10.chm906072
ms.prod: excel
api_name:
- Excel.SlicerItem
ms.assetid: cb93cd82-fc3a-f6b7-ae64-db6312db649d
ms.date: 04/02/2019
localization_priority: Normal
---


# SlicerItem object (Excel)

Represents an item in a slicer.


## Remarks

To access the **SlicerItem** object that represents the currently selected button in the slicer, use the **[ActiveItem](Excel.Slicer.ActiveItem.md)** property of the **Slicer** object. 

To access the **[SlicerItems](Excel.SlicerItems.md)** collection that represents all the items in a slicer filtering a PivotTable, use the **[SlicerItems](Excel.SlicerCache.SlicerItems.md)** property of the **SlicerCache** object that is associated with the **Slicer** object. 

To access the **SlicerItems** collection that represents the items in a slicer filtering a level of an OLAP hierarchy, use the **[SlicerItems](Excel.SlicerCacheLevel.SlicerItems.md)** property of the **SlicerCacheLevel** object that represents that level of the hierarchy.


## Properties

- [Application](Excel.SlicerItem.Application.md)
- [Caption](Excel.SlicerItem.Caption.md)
- [Creator](Excel.SlicerItem.Creator.md)
- [HasData](Excel.SlicerItem.HasData.md)
- [Name](Excel.SlicerItem.Name.md)
- [Parent](Excel.SlicerItem.Parent.md)
- [Selected](Excel.SlicerItem.Selected.md)
- [SourceName](Excel.SlicerItem.SourceName.md)
- [SourceNameStandard](Excel.SlicerItem.SourceNameStandard.md)
- [Value](Excel.SlicerItem.Value.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
