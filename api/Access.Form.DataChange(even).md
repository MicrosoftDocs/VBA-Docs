---
title: Form.DataChange event (Access)
keywords: vbaac10.chm13685
f1_keywords:
- vbaac10.chm13685
ms.prod: access
api_name:
- Access.Form.DataChange
ms.assetid: 026fddb4-2a43-095c-9460-98c12378735c
ms.date: 03/08/2019
localization_priority: Normal
---


# Form.DataChange event (Access)

Occurs when certain properties are changed or when certain methods are executed in the specified PivotTable view.


## Syntax

_expression_.**DataChange** (_Reason_)

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Reason_|Required|**Long**|A **PivotDataReasonEnum** constant that indicates the reason that this event was triggered.|

## Return value

Nothing


## Example

The following example demonstrates the syntax for a subroutine that traps the **DataChange** event.

```vb
Private Sub Form_DataChange(Reason As Long) 
 If Reason = OWC.plDataReasonDisplayCellColorChange Then 
 MsgBox "The cell display color was changed." 
 End If 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
