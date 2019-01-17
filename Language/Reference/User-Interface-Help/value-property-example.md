---
title: Value property example
keywords: fm20.chm5225132
f1_keywords:
- fm20.chm5225132
ms.prod: office
ms.assetid: 7d98bbfa-9f19-b554-b327-554b12508b70
ms.date: 11/14/2018
localization_priority: Normal
---


# Value property example

The following example demonstrates the values that the different types of controls can have by displaying the **[Value](value-property-microsoft-forms.md)** property of a selected control. 

The user chooses a control by pressing Tab or by clicking the control. Depending on the type of control, the user can also specify a value for the control by typing in the text area of the control, by clicking one or more times on the control, or by selecting an item, page, or tab within the control. The user can display the value of the selected control by clicking the appropriately labeled **[CommandButton](commandbutton-control.md)**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **CommandButton** named CommandButton1.   
- A **[TextBox](textbox-control.md)** named TextBox1.    
- A **[CheckBox](checkbox-control.md)** named CheckBox1.   
- A **[ComboBox](combobox-control.md)** named ComboBox1.    
- A **CommandButton** named CommandButton2.   
- A **[ListBox](listbox-control.md)** named ListBox1.   
- A **[MultiPage](multipage-control.md)** named MultiPage1.   
- Two **[OptionButton](optionbutton-control.md)** controls named OptionButton1 and OptionButton2.   
- A **[ScrollBar](scrollbar-control.md)** named ScrollBar1.   
- A **[SpinButton](spinbutton-control.md)** named SpinButton1.    
- A **[TabStrip](tabstrip-control.md)** named TabStrip1.   
- A **TextBox** named TextBox2.   
- A **[ToggleButton](togglebutton-control.md)** named ToggleButton1.
    

```vb
Dim i As Integer 
 
Private Sub CommandButton1_Click() 
 TextBox1.Text = "Value of " & ActiveControl.Name _ 
 & " is " & ActiveControl.Value 
End Sub 
 
Private Sub UserForm_Initialize() 
 CommandButton1.Caption = "Get value of " _ 
 & "current control" 
 CommandButton1.AutoSize = True 
 CommandButton1.TakeFocusOnClick = False 
 CommandButton1.TabStop = False 
 
 TextBox1.AutoSize = True 
 
 For i = 0 To 10 
 ComboBox1.AddItem "Choice " & (i + 1) 
 ListBox1.AddItem "Selection " & (100 - i) 
 Next i 
 
 CheckBox1.TripleState = True 
 ToggleButton1.TripleState = True 
 
 TextBox2.Text = "Enter text here." 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]