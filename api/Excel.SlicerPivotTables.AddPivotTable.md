---
title: SlicerPivotTables.AddPivotTable method (Excel)
keywords: vbaxl10.chm911077
f1_keywords:
- vbaxl10.chm911077
ms.prod: excel
api_name:
- Excel.SlicerPivotTables.AddPivotTable
ms.assetid: c5fc95c6-0fb9-1c8f-5b12-8a4c0f9f10c7
ms.date: 05/16/2019
localization_priority: Normal
---


# SlicerPivotTables.AddPivotTable method (Excel)

Adds a reference to a PivotTable to the **SlicerPivotTables** collection.


## Syntax

_expression_.**AddPivotTable** (_PivotTable_)

_expression_ A variable that represents a **[SlicerPivotTables](Excel.SlicerPivotTables.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PivotTable_|Required| **PivotTable**|A **[PivotTable](Excel.PivotTable.md)** object that represents the PivotTable to add.|

## Return value

Nothing


## Remarks

When a PivotTable is added to the **SlicerPivotTables** collection, it can be filtered by its parent **[SlicerCache](Excel.SlicerCache.md)** and the slicers associated with it.


## Example

The following code example adds PivotTable1 to the slicer cache associated with the Customer slicer.

```vb
Dim pvts As SlicerPivotTables 
Set pvts = ActiveWorkbook.SlicerCaches("Slicer_Customer").PivotTables 
pvts.AddPivotTable(ActiveSheet.PivotTables("PivotTable1"))
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]