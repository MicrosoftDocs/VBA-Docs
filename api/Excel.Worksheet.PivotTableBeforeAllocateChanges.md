---
title: Worksheet.PivotTableBeforeAllocateChanges event (Excel)
keywords: vbaxl10.chm502083
f1_keywords:
- vbaxl10.chm502083
ms.prod: excel
api_name:
- Excel.Worksheet.PivotTableBeforeAllocateChanges
ms.assetid: 220729d9-2da4-53fb-2910-26cc8f835da7
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.PivotTableBeforeAllocateChanges event (Excel)

Occurs before changes are applied to a PivotTable.


## Syntax

_expression_.**PivotTableBeforeAllocateChanges** (_TargetPivotTable_, _ValueChangeStart_, _ValueChangeEnd_, _Cancel_)

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _TargetPivotTable_|Required| **[PivotTable](Excel.PivotTable.md)**|The PivotTable that contains the changes to apply.|
| _ValueChangeStart_|Required| **Long**|The index to the first change in the associated **[PivotTableChangeList](Excel.PivotTableChangeList.md)** collection. The index is specified by the **[Order](Excel.ValueChange.Order.md)** property of the **ValueChange** object in the **PivotTableChangeList** collection.|
| _ValueChangeEnd_|Required| **Long**|The index to the last change in the associated **PivotTableChangeList** collection. The index is specified by the **Order** property of the **ValueChange** object in the **PivotTableChangeList** collection.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the changes are not applied to the PivotTable and all edits are lost.|

## Return value

**Nothing**


## Remarks

The **PivotTableBeforeAllocateChanges** event occurs immediately before Excel executes an **UPDATE CUBE** statement to apply all changes to the PivotTable's OLAP data source, and immediately after the user has chosen to apply changes in the user interface.


## Example

The following code example prompts the user before updates are applied to the PivotTable's OLAP data source.

```vb
Sub Worksheet_PivotTableBeforeAllocateChanges(ByVal TargetPivotTable As PivotTable, _ 
 ByVal ValueChangeStart As Long, ByVal ValueChangeEnd As Long, Cancel As Boolean) 
 Dim UserChoice As VbMsgBoxResult 
 
 UserChoice = MsgBox("Allow updates to be applied to: " + TargetPivotTable.Name + "?", vbYesNo) 
 If UserChoice = vbNo Then Cancel = True 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]