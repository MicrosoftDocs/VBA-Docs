---
title: MatchEntry property, OptionButton control example
keywords: fm20.chm5225120
f1_keywords:
- fm20.chm5225120
ms.prod: office
ms.assetid: c68bae6a-b2cc-8616-bffb-9b7369fd9749
ms.date: 11/14/2018
localization_priority: Normal
---


# MatchEntry property, OptionButton control example

The following example uses the **[MatchEntry](matchentry-property.md)** property to demonstrate character matching that is available for **[ComboBox](combobox-control.md)** and **[ListBox](listbox-control.md)**. 

In this example, the user can set the type of matching with the **[OptionButton](optionbutton-control.md)** controls, and then type into the **ComboBox** to specify an item from its list.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- Three **OptionButton** controls named OptionButton1 through OptionButton3.    
- A **ComboBox** named ComboBox1.
    

```vb
Private Sub OptionButton1_Click() 
 ComboBox1.MatchEntry = fmMatchEntryNone 
End Sub 
 
Private Sub OptionButton2_Click() 
 ComboBox1.MatchEntry = fmMatchEntryFirstLetter 
End Sub 
 
Private Sub OptionButton3_Click() 
 ComboBox1.MatchEntry = fmMatchEntryComplete 
End Sub 
 
Private Sub UserForm_Initialize() 
 Dim i As Integer 
 
 For i = 1 To 9 
 ComboBox1.AddItem "Choice " & i 
 Next i 
 ComboBox1.AddItem "Chocoholic" 
 
 
 OptionButton1.Caption = "No matching" 
 OptionButton1.Value = True 
 
 OptionButton2.Caption = "Basic matching" 
 OptionButton3.Caption = "Extended matching" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]