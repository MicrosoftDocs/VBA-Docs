---
title: Workbook.SheetPivotTableBeforeAllocateChanges event (Excel)
keywords: vbaxl10.chm503103
f1_keywords:
- vbaxl10.chm503103
ms.prod: excel
api_name:
- Excel.Workbook.SheetPivotTableBeforeAllocateChanges
ms.assetid: 2f767b5b-27fb-33de-c91d-76bbc52ea171
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.SheetPivotTableBeforeAllocateChanges event (Excel)

Occurs before changes are applied to a PivotTable.


## Syntax

_expression_.**SheetPivotTableBeforeAllocateChanges** (_Sh_, _TargetPivotTable_, _ValueChangeStart_, _ValueChangeEnd_, _Cancel_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The worksheet that contains the PivotTable.|
| _TargetPivotTable_|Required| **[PivotTable](Excel.PivotTable.md)**|The PivotTable that contains the changes to apply.|
| _ValueChangeStart_|Required| **Long**|The index to the first change in the associated **[PivotTableChangeList](Excel.PivotTableChangeList.md)** collection. The index is specified by the **[Order](Excel.ValueChange.Order.md)** property of the **ValueChange** object in the **PivotTableChangeList** collection.|
| _ValueChangeEnd_|Required| **Long**|The index to the last change in the associated **PivotTableChangeList** collection. The index is specified by the **Order** property of the **ValueChange** object in the **PivotTableChangeList** collection.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the changes are not applied to the PivotTable, and all edits are lost.|

## Return value

**Nothing**


## Remarks

The **SheetPivotTableBeforeAllocateChanges** event occurs immediately before Excel executes an **UPDATE CUBE** statement to apply all changes to the PivotTable's OLAP data source, and immediately after the user has chosen to apply changes in the user interface.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]