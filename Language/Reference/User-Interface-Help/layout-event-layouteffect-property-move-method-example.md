---
title: Layout event, LayoutEffect property, Move method example
keywords: fm20.chm5225128
f1_keywords:
- fm20.chm5225128
ms.prod: office
ms.assetid: c3585b29-d100-89a8-8e64-3afe5dbae8b2
ms.date: 11/14/2018
localization_priority: Normal
---


# Layout event, LayoutEffect property, Move method example

The following example moves a selected control on a form with the **[Move](move-method.md)** method, and uses the **[Layout](layout-event.md)** event and **[LayoutEffect](layouteffect-property.md)** property to identify the control that moved (and changed the layout of the **[UserForm](userform-window.md)**). 

The user clicks a control to move and then clicks the **[CommandButton](commandbutton-control.md)**. A message box displays the name of the control that is moving.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **[TextBox](textbox-control.md)** named TextBox1.    
- A **[ComboBox](combobox-control.md)** named ComboBox1.    
- An **[OptionButton](optionbutton-control.md)** named OptionButton1.    
- A **CommandButton** named CommandButton1.    
- A **[ToggleButton](togglebutton-control.md)** named ToggleButton1.
    

```vb
Private Sub UserForm_Initialize() 
 CommandButton1.Caption = "Move current control" 
 CommandButton1.AutoSize = True 
 CommandButton1.TakeFocusOnClick = False 
 
 ToggleButton1.Caption = "Use Layout Event" 
 ToggleButton1.Value = True 
End Sub 
 
Private Sub CommandButton1_Click() 
 If ActiveControl.Name = "ToggleButton1" Then 
 'Keep it stationary 
 Else 
 'Move the control, using Layout event when 
 'ToggleButton1.Value is True 
 ActiveControl.Move 0, 0, , , _ 
 ToggleButton1.Value 
 End If 
End Sub 
 
Private Sub UserForm_Layout() 
 Dim MyControl As Control 
 
 MsgBox "In the Layout Event" 
 
 'Find the control that is moving. 
 For Each MyControl In Controls 
 If MyControl.LayoutEffect = _ 
 fmLayoutEffectInitiate Then 
 MsgBox MyControl.Name & " is moving." 
 Exit For 
 End If 
 Next 
End Sub 
 
Private Sub ToggleButton1_Click() 
 If ToggleButton1.Value = True Then 
 ToggleButton1.Caption = "Use Layout Event" 
 Else 
 ToggleButton1.Caption = "No Layout Event" 
 End If 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]