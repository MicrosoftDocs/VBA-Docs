---
title: Workbook.SheetPivotTableBeforeCommitChanges event (Excel)
keywords: vbaxl10.chm503104
f1_keywords:
- vbaxl10.chm503104
ms.prod: excel
api_name:
- Excel.Workbook.SheetPivotTableBeforeCommitChanges
ms.assetid: 7e189a4f-a349-f862-375a-fa66311629cc
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.SheetPivotTableBeforeCommitChanges event (Excel)

Occurs before changes are committed against the OLAP data source for a PivotTable.


## Syntax

_expression_.**SheetPivotTableBeforeCommitChanges** (_Sh_, _TargetPivotTable_, _ValueChangeStart_, _ValueChangeEnd_, _Cancel_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The worksheet that contains the PivotTable.|
| _TargetPivotTable_|Required| **[PivotTable](Excel.PivotTable.md)**|The PivotTable that contains the changes to commit.|
| _ValueChangeStart_|Required| **Long**|The index to the first change in the associated **[PivotTableChangeList](Excel.PivotTableChangeList.md)** object. The index is specified by the **[Order](Excel.ValueChange.Order.md)** property of the **ValueChange** object in the **PivotTableChangeList** collection.|
| _ValueChangeEnd_|Required| **Long**|The index to the last change in the associated **PivotTableChangeList** object. The index is specified by the **Order** property of the **ValueChange** object in the **PivotTableChangeList** collection.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the changes are not committed against the OLAP data source of the PivotTable.|

## Return value

**Nothing**


## Remarks

The **SheetPivotTableBeforeCommitChanges** event occurs immediately before Excel executes a **COMMIT TRANSACTION** statement against the PivotTable's OLAP data source, and immediately after the user has chosen to save changes for the whole PivotTable.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]