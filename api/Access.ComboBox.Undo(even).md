---
title: ComboBox.Undo event (Access)
keywords: vbaac10.chm14228
f1_keywords:
- vbaac10.chm14228
ms.prod: access
api_name:
- Access.ComboBox.Undo
ms.assetid: d1064051-bbf9-ce00-c43e-19775879185c
ms.date: 02/28/2019
localization_priority: Normal
---


# ComboBox.Undo event (Access)

Occurs when the user undoes a change.


## Syntax

_expression_.**Undo** (_Cancel_)

_expression_ A variable that represents a **[ComboBox](Access.ComboBox.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required|**Integer**|Set this argument to **True** to cancel the undo operation and leave the control or form in its edited state.|

## Remarks

The **Undo** event for controls occurs whenever the user returns a control to its original state by choosing the **Undo Field/Record** button on the command bar, choosing the **Undo** button, pressing the Esc key, or calling the **Undo** method of the specified control. The control needs to have focus in all three cases. The event does not occur if the user chooses the **Undo Typing** button on the command bar.

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