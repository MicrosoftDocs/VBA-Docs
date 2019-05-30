---
title: Worksheet.PivotTableBeforeDiscardChanges event (Excel)
keywords: vbaxl10.chm502085
f1_keywords:
- vbaxl10.chm502085
ms.prod: excel
api_name:
- Excel.Worksheet.PivotTableBeforeDiscardChanges
ms.assetid: 94a480fa-ce06-e7d7-d4b4-ac21be0525ac
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.PivotTableBeforeDiscardChanges event (Excel)

Occurs before changes to a PivotTable are discarded.


## Syntax

_expression_.**PivotTableBeforeDiscardChanges** (_TargetPivotTable_, _ValueChangeStart_, _ValueChangeEnd_)

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _TargetPivotTable_|Required| **[PivotTable](Excel.PivotTable.md)**|The PivotTable that contains the changes to discard.|
| _ValueChangeStart_|Required| **Long**|The index to the first change in the associated **[PivotTableChangeList](Excel.PivotTableChangeList.md)** object. The index is specified by the **[Order](Excel.ValueChange.Order.md)** property of the **ValueChange** object in the **PivotTableChangeList** collection.|
| _ValueChangeEnd_|Required| **Long**|The index to the last change in the associated **PivotTableChangeList** object. The index is specified by the **Order** property of the **ValueChange** object in the **PivotTableChangeList** collection.|

## Return value

**Nothing**


## Remarks

Occurs immediately before Excel executes a **ROLLBACK TRANSACTION** statement against the OLAP data source, if a transaction is still active, and then discards all edited values in the PivotTable, after the user has chosen to discard changes.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]