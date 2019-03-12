---
title: Tag property example
keywords: fm20.chm5225129
f1_keywords:
- fm20.chm5225129
ms.prod: office
ms.assetid: 23ace8e6-5d8a-6b61-d69d-eb403be6e605
ms.date: 11/14/2018
localization_priority: Normal
---


# Tag property example

The following example uses the **[Tag](tag-property.md)** property to store additional information about each control on the **[UserForm](userform-window.md)**. The user clicks a control and then clicks the **[CommandButton](commandbutton-control.md)**. The contents of **Tag** for the appropriate control are returned in the **[TextBox](textbox-control.md)**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **TextBox** named TextBox1.    
- A **CommandButton** named CommandButton1.    
- A **[ScrollBar](scrollbar-control.md)** named ScrollBar1.    
- A **[ComboBox](combobox-control.md)** named ComboBox1.    
- A **[MultiPage](multipage-control.md)** named MultiPage1.
    

```vb
Private Sub CommandButton1_Click() 
 TextBox1.Text = ActiveControl.Tag 
End Sub 
 
Private Sub UserForm_Initialize() 
 TextBox1.Locked = True 
 TextBox1.Tag = "Display area for Tag properties." 
 TextBox1.AutoSize = True 
 
 CommandButton1.Caption = "Show Tag of Current " _ 
 & "Control." 
 CommandButton1.AutoSize = True 
 CommandButton1.WordWrap = True 
 CommandButton1.TakeFocusOnClick = False 
 CommandButton1.Tag = "Shows tag of control " _ 
 & "that has the focus." 
 
 ComboBox1.Style = fmStyleDropDownList 
 ComboBox1.Tag = "ComboBox Style is that of " _ 
 & "a ListBox." 
 
 ScrollBar1.Max = 100 
 ScrollBar1.Min = -273 
 ScrollBar1.Tag = "Max = " & ScrollBar1.Max _ 
 & " , Min = " & ScrollBar1.Min 
 
 MultiPage1.Pages.Add 
 MultiPage1.Pages.Add 
 MultiPage1.Tag = "This MultiPage has " _ 
 & MultiPage1.Pages.Count & " pages." 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
