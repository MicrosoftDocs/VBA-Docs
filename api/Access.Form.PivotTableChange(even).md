---
title: Form.PivotTableChange event (Access)
keywords: vbaac10.chm13669
f1_keywords:
- vbaac10.chm13669
ms.prod: access
api_name:
- Access.Form.PivotTableChange
ms.assetid: 8b4a8c9a-c8a3-648d-968d-edcb7cb94956
ms.date: 03/08/2019
localization_priority: Normal
---


# Form.PivotTableChange event (Access)

Occurs whenever the specified PivotTable view field, field set, or total is added or deleted.


## Syntax

_expression_.**PivotTableChange** (_Reason_)

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Reason_|Required|**Long**|A **PivotTableReasonEnum** constant that indicates how the PivotTable list changed.|

## Example

The following example demonstrates the syntax for a subroutine that traps the **PivotTableChange** event.

```vb
Private Sub Form_PivotTableChange(Reason As Long) 
 Select Case Reason 
 Case OWC.plPivotTableReasonTotalAdded 
 MsgBox "A total was added!" 
 Case OWC.plPivotTableReasonFieldSetAdded 
 MsgBox "A field set was added!" 
 Case OWC.plPivotTableReasonFieldAdded 
 MsgBox "A field was added!" 
 End Select 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]