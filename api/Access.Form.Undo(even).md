---
title: Form.Undo event (Access)
keywords: vbaac10.chm13663
f1_keywords:
- vbaac10.chm13663
ms.prod: access
api_name:
- Access.Form.Undo
ms.assetid: fdcf98c1-c560-1c29-586d-6c4eb4a6ccd0
ms.date: 02/28/2019
localization_priority: Normal
---


# Form.Undo event (Access)

Occurs when the user undoes a change.


## Syntax

_expression_.**Undo** (_Cancel_)

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required|**Integer**|Set this argument to **True** to cancel the undo operation and leave the control or form in its edited state.|

## Remarks

The **Undo** event for controls occurs whenever the user returns a control to its original state by choosing the **Undo Field/Record** button on the command bar, choosing the **Undo** button, pressing the Esc key, or calling the **Undo** method of the specified control. The control needs to have focus in all three cases. The event does not occur if the user chooses the **Undo Typing** button on the command bar.

The **Undo** event for forms occurs whenever the user returns a form to its original state by choosing the **Undo** button, pressing the Esc key, or calling the **Undo** method of the specified form.


## Example

The following example demonstrates the syntax for a subroutine that traps the **Undo** event for a form.

```vb
Private Sub Form_Undo(Cancel As Integer) 
 Dim intResponse As Integer 
 Dim strPrompt As String 
 
 strPrompt = "Cancel the undo operation?" 
 
 intResponse = MsgBox(strPrompt, vbYesNo) 
 
 If intResponse = vbYes Then 
 Cancel = True 
 Else 
 Cancel = False 
 End If 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]