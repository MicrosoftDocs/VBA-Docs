---
title: Form.ViewChange event (Access)
keywords: vbaac10.chm13684
f1_keywords:
- vbaac10.chm13684
ms.prod: access
api_name:
- Access.Form.ViewChange
ms.assetid: a3788eca-783f-cb5d-1a7b-1c4a23648629
ms.date: 03/08/2019
localization_priority: Normal
---


# Form.ViewChange event (Access)

Occurs whenever the specified PivotChart view or PivotTable view is redrawn.


## Syntax

_expression_.**ViewChange** (_Reason_)

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Reason_|Required|**Long**| A **PivotViewReasonEnum** constant that indicates how the view was changed. _Reason_ always returns 1 for PivotChart views.|

## Example

The following example demonstrates the syntax for a subroutine that traps the **ViewChange** event.

```vb
Private Sub Form_ViewChange(ByVal Reason As Long) 
 If Reason = OWC.plViewReasonShowDetails Then 
 MsgBox "You've opted to show details." 
 End If 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]