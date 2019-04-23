---
title: TakeFocusOnClick property example
keywords: fm20.chm5225119
f1_keywords:
- fm20.chm5225119
ms.prod: office
ms.assetid: fdc5a590-eee9-0ab2-aead-f3c02abf0eab
ms.date: 11/14/2018
localization_priority: Normal
---


# TakeFocusOnClick property example

The following example uses the **[TakeFocusOnClick](takefocusonclick-property.md)** property to control whether a **[CommandButton](commandbutton-control.md)** receives the focus when the user clicks it. 

The user clicks a control other than CommandButton1 and then clicks CommandButton1. If **TakeFocusOnClick** is **True**, CommandButton1 receives the focus after it is clicked. The user can change the value of **TakeFocusOnClick** by clicking the **[ToggleButton](togglebutton-control.md)**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **CommandButton** named CommandButton1.   
- A **ToggleButton** named ToggleButton1.   
- One or two other controls, such as an **[OptionButton](optionbutton-control.md)** or **[ListBox](listbox-control.md)**.
    

```vb
Private Sub CommandButton1_Click() 
 MsgBox "Watch CommandButton1 to see if it " _ 
 & "takes the focus." 
End Sub 
 
Private Sub ToggleButton1_Click() 
 If ToggleButton1 = True Then 
 CommandButton1.TakeFocusOnClick = True 
 ToggleButton1.Caption = "TakeFocusOnClick On" 
 Else 
 CommandButton1.TakeFocusOnClick = False 
 ToggleButton1.Caption = "TakeFocusOnClick Off" 
 End If 
End Sub 
 
Private Sub UserForm_Initialize() 
 CommandButton1.Caption = "Show Message" 
 
 ToggleButton1.Caption = "TakeFocusOnClick On" 
 ToggleButton1.Value = True 
 ToggleButton1.Width = 90 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]