---
title: Worksheet.PivotTableBeforeCommitChanges event (Excel)
keywords: vbaxl10.chm502084
f1_keywords:
- vbaxl10.chm502084
ms.prod: excel
api_name:
- Excel.Worksheet.PivotTableBeforeCommitChanges
ms.assetid: 4dfcfd60-9249-4eed-1bb3-183b5c567125
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.PivotTableBeforeCommitChanges event (Excel)

Occurs before changes are committed against the OLAP data source for a PivotTable.


## Syntax

_expression_.**PivotTableBeforeCommitChanges** (_TargetPivotTable_, _ValueChangeStart_, _ValueChangeEnd_, _Cancel_)

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _TargetPivotTable_|Required| **[PivotTable](Excel.PivotTable.md)**|The PivotTable that contains the changes to commit.|
| _ValueChangeStart_|Required| **Long**|The index to the first change in the associated **[PivotTableChangeList](Excel.PivotTableChangeList.md)** object. The index is specified by the **[Order](Excel.ValueChange.Order.md)** property of the **ValueChange** object in the **PivotTableChangeList** collection.|
| _ValueChangeEnd_|Required| **Long**|The index to the last change in the associated **PivotTableChangeList** object. The index is specified by the **Order** property of the **ValueChange** object in the **PivotTableChangeList** collection.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the changes are not committed against the OLAP data source of the PivotTable.|

## Return value

**Nothing**


## Remarks

The **PivotTableBeforeCommitChanges** event occurs immediately before Excel executes a **COMMIT TRANSACTION** statement against the PivotTable's OLAP data source, and immediately after the user has chosen to save changes for the whole PivotTable.


## Example

The following code example prompts the user before changes are committed to the PivotTable's OLAP data source.

```vb
Sub Worksheet_PivotTableBeforeCommitChanges(ByVal TargetPivotTable As PivotTable, _ 
 ByVal ValueChangeStart As Long, ByVal ValueChangeEnd As Long, Cancel As Boolean) 
 
 Dim UserChoice As VbMsgBoxResult 
 
 UserChoice = MsgBox("Allow updates to be saved to: " + TargetPivotTable.Name + "?", vbYesNo) 
 If UserChoice = vbNo Then Cancel = True 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]