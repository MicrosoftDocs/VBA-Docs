---
title: Enter, Exit events, ActiveControl property example
keywords: fm20.chm5225152
f1_keywords:
- fm20.chm5225152
ms.prod: office
ms.assetid: 8d3123e3-e5b1-cb8f-0f89-de308c3eecda
ms.date: 11/14/2018
localization_priority: Normal
---


# Enter, Exit events, ActiveControl property example

The following example uses the **[ActiveControl](activecontrol-property.md)** property in a subroutine that tracks the controls a user visits. The **[Enter](enter-exit-events.md)** event for each control calls the TraceFocus subroutine to identify the control that has the focus.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains the following controls:

- A **[ScrollBar](scrollbar-control.md)** named ScrollBar1.   
- A **[ListBox](listbox-control.md)** named ListBox1.    
- Two **[OptionButton](optionbutton-control.md)** controls named OptionButton1 and OptionButton2.    
- A **[Frame](frame-control.md)** named Frame1.
    

```vb
Dim MyControl As Control 
 
Private Sub TraceFocus() 
 ListBox1.AddItem ActiveControl.Name 
 ListBox1.List(ListBox1.ListCount - 1, 1) = _ 
 ActiveControl.TabIndex 
End Sub 
 
Private Sub UserForm_Initialize() 
 ListBox1.ColumnCount = 2 
 ListBox1.AddItem "Controls Visited" 
 ListBox1.List(0, 1) = "Control Index" 
End Sub 
 
Private Sub Frame1_Enter() 
 TraceFocus 
End Sub 
 
Private Sub ListBox1_Enter() 
 TraceFocus 
End Sub 
 
Private Sub OptionButton1_Enter() 
 TraceFocus 
End Sub 
 
Private Sub OptionButton2_Enter() 
 TraceFocus 
End Sub 
 
Private Sub ScrollBar1_Enter() 
 TraceFocus 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]