---
title: TabStop property example
keywords: fm20.chm5225118
f1_keywords:
- fm20.chm5225118
ms.prod: office
ms.assetid: 120e875d-0dff-6b69-31e6-60da49d3be84
ms.date: 11/14/2018
localization_priority: Normal
---


# TabStop property example

The following example uses the **[TabStop](tabstop-property-vbe-dev.md)** property to control whether a user can press Tab to move the focus to a particular control. The user presses Tab to move the focus among the controls on the form, and then clicks the **[ToggleButton](togglebutton-control.md)** to change **TabStop** for CommandButton1. When **TabStop** is **False**, CommandButton1 will not receive the focus by using Tab.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **[CommandButton](commandbutton-control.md)** named CommandButton1.    
- A **ToggleButton** named ToggleButton1.    
- One or two other controls, such as an **[OptionButton](optionbutton-control.md)** or **[ListBox](listbox-control.md)**.
    

```vb
Private Sub CommandButton1_Click() 
 MsgBox "Clicked CommandButton1." 
End Sub 
 
Private Sub ToggleButton1_Click() 
 If ToggleButton1 = True Then 
 CommandButton1.TabStop = True 
 ToggleButton1.Caption = "TabStop On" 
 Else 
 CommandButton1.TabStop = False 
 ToggleButton1.Caption = "TabStop Off" 
 End If 
End Sub 
 
Private Sub UserForm_Initialize() 
 CommandButton1.Caption = "Show Message" 
 
 ToggleButton1.Caption = "TabStop On" 
 ToggleButton1.Value = True 
 ToggleButton1.Width = 90 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]