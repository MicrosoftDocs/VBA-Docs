---
title: Application.SheetPivotTableBeforeDiscardChanges event (Excel)
keywords: vbaxl10.chm504107
f1_keywords:
- vbaxl10.chm504107
ms.prod: excel
api_name:
- Excel.Application.SheetPivotTableBeforeDiscardChanges
ms.assetid: 8623adc6-d256-bebb-fe35-8710390af19f
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.SheetPivotTableBeforeDiscardChanges event (Excel)

Occurs before changes to a PivotTable are discarded.


## Syntax

_expression_.**SheetPivotTableBeforeDiscardChanges** (_Sh_, _TargetPivotTable_, _ValueChangeStart_, _ValueChangeEnd_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**||
| _TargetPivotTable_|Required| **[PivotTable](Excel.PivotTable.md)**|The PivotTable that contains the changes to discard.|
| _ValueChangeStart_|Required| **Long**|The index to the first change in the associated **[PivotTableChangeList](Excel.PivotTableChangeList.md)** object. The index is specified by the **[Order](Excel.ValueChange.Order.md)** property of the **ValueChange** object in the **PivotTableChangeList** collection.|
| _ValueChangeEnd_|Required| **Long**|The index to the last change in the associated **PivotTableChangeList** object. The index is specified by the **Order** property of the **ValueChange** object in the **PivotTableChangeList** collection.|

## Return value

Nothing


## Remarks

Occurs immediately before Excel executes a **ROLLBACK TRANSACTION** statement against the OLAP data source, if a transaction is still active, and then discards all edited values in the PivotTable after the user has chosen to discard changes.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]