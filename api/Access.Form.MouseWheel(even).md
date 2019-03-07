---
title: Form.MouseWheel event (Access)
keywords: vbaac10.chm13683
f1_keywords:
- vbaac10.chm13683
ms.prod: access
api_name:
- Access.Form.MouseWheel
ms.assetid: eec18d43-1cee-463c-37e6-760eccb0b890
ms.date: 03/08/2019
localization_priority: Normal
---


# Form.MouseWheel event (Access)

Occurs when the user rolls the mouse wheel in Form view, Split Form view, Datasheet view, Layout view, PivotChart view, or PivotTable view.


## Syntax

_expression_.**MouseWheel** (_Page_, _Count_)

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Page_|Required|**Boolean**|**True** if the page was changed.|
| _Count_|Required|**Long**|The number of lines by which the view was scrolled with the mouse wheel.|

## Example

The following example demonstrates the syntax for a subroutine that traps the **MouseWheel** event.

```vb
Private Sub Form_MouseWheel( _ 
 ByVal Page As Boolean, ByVal Count As Long) 
 If Page = True Then 
 MsgBox "You've moved to another page." 
 End If 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]