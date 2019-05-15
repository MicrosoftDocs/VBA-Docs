---
title: SlicerPivotTables.RemovePivotTable method (Excel)
keywords: vbaxl10.chm911078
f1_keywords:
- vbaxl10.chm911078
ms.prod: excel
api_name:
- Excel.SlicerPivotTables.RemovePivotTable
ms.assetid: ebc4cc53-c406-3ae4-06e7-094a1ba32af2
ms.date: 05/16/2019
localization_priority: Normal
---


# SlicerPivotTables.RemovePivotTable method (Excel)

Removes a reference to a PivotTable from the **SlicerPivotTables** collection.


## Syntax

_expression_.**RemovePivotTable** (_PivotTable_)

_expression_ A variable that represents a **[SlicerPivotTables](Excel.SlicerPivotTables.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PivotTable_|Required| **Variant**|A **[PivotTable](Excel.PivotTable.md)** object that represents the PivotTable to remove, or the name or index of the PivotTable in the collection.|

## Return value

Nothing


## Remarks

When a PivotTable is removed from the **SlicerPivotTables** collection, it is no longer filtered by its parent **[SlicerCache](Excel.SlicerCache.md)** object and the slicers associated with it.


## Example

The following code example removes PivotTable1 from the slicer cache associated with the Customer slicer.

```vb
Dim pvts As SlicerPivotTables 
Set pvts = ActiveWorkbook.SlicerCaches("Slicer_Customer").PivotTables 
pvts.RemovePivotTable("PivotTable1")
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]