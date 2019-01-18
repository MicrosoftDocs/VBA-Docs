---
title: SpinDown, SpinUp events, Delay property example
keywords: fm20.chm5225147
f1_keywords:
- fm20.chm5225147
ms.prod: office
ms.assetid: a7c32938-d1b3-9962-8333-716ab8b09337
ms.date: 11/14/2018
localization_priority: Normal
---


# SpinDown, SpinUp events, Delay property example

The following example demonstrates the time interval between successive **[Change](change-event.md)**, **[SpinUp and SpinDown](spindown-spinup-events.md)** events that occur when a user holds down the mouse button to change the value of a **[SpinButton](spinbutton-control.md)** or **[ScrollBar](scrollbar-control.md)**.

In this example, the user chooses a delay setting, and then clicks and holds down either side of a **SpinButton**. The **SpinUp** and **SpinDown** events are recorded in a **[ListBox](listbox-control.md)** as they are initiated.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **SpinButton** named SpinButton1.    
- Two **[OptionButton](optionbutton-control.md)** controls named OptionButton1 and OptionButton2.    
- A **ListBox** named ListBox1.
    

```vb
Dim EventCount As Long 
 
Private Sub ResetControl() 
 ListBox1.Clear 
 EventCount = 0 
 SpinButton1.Value = 5000 
End Sub 
 
Private Sub UserForm_Initialize() 
 SpinButton1.Min = 0 
 SpinButton1.Max = 10000 
 ResetControl 
 
 SpinButton1.Delay = 50 
 OptionButton1.Caption = "50 millisecond delay" 
 OptionButton2.Caption = "250 millisecond delay" 
 
 OptionButton1.Value = True 
End Sub 
 
Private Sub OptionButton1_Click() 
 SpinButton1.Delay = 50 
 ResetControl 
End Sub 
 
Private Sub OptionButton2_Click() 
 SpinButton1.Delay = 250 
 ResetControl 
End Sub 
 
Private Sub SpinButton1_SpinDown() 
 EventCount = EventCount + 1 
 ListBox1.AddItem EventCount 
End Sub 
 
Private Sub SpinButton1_SpinUp() 
 EventCount = EventCount + 1 
 ListBox1.AddItem EventCount 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]