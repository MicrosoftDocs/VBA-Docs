---
title: ListStyle, MultiSelect properties example
keywords: fm20.chm5225142
f1_keywords:
- fm20.chm5225142
ms.prod: office
ms.assetid: 8a5ea21b-fadb-994c-6df8-e40e29094f42
ms.date: 11/14/2018
localization_priority: Normal
---


# ListStyle, MultiSelect properties example

The following example uses the **[ListStyle](liststyle-property.md)** and **[MultiSelect](multiselect-property-microsoft-forms.md)** properties to control the appearance of a **[ListBox](listbox-control.md)**. 

The user chooses a value for **ListStyle** using the **[ToggleButton](togglebutton-control.md)** and chooses an **[OptionButton](optionbutton-control.md)** for one of the **MultiSelect** values. The appearance of the **ListBox** changes accordingly, as well as the selection behavior within the **ListBox**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **ListBox** named ListBox1.    
- A **[Label](label-control.md)** named Label1.    
- Three **OptionButton** controls named OptionButton1 through OptionButton3.   
- A **ToggleButton** named ToggleButton1.
    

```vb
Private Sub UserForm_Initialize() 
 Dim i As Integer 
 
 For i = 1 To 8 
 ListBox1.AddItem "Choice" & (ListBox1.ListCount + 1) 
 Next i 
 
 Label1.Caption = "MultiSelect Choices" 
 Label1.AutoSize = True 
 
 ListBox1.MultiSelect = fmMultiSelectSingle 
 OptionButton1.Caption = "Single entry" 
 OptionButton1.Value = True 
 OptionButton2.Caption = "Multiple entries" 
 OptionButton3.Caption = "Extended entries" 
 
 ToggleButton1.Caption = "ListStyle - Plain" 
 ToggleButton1.Value = True 
 ToggleButton1.Width = 90 
 ToggleButton1.Height = 30 
End Sub 
 
Private Sub OptionButton1_Click() 
 ListBox1.MultiSelect = fmMultiSelectSingle 
End Sub 
 
Private Sub OptionButton2_Click() 
 ListBox1.MultiSelect = fmMultiSelectMulti 
End Sub 
 
Private Sub OptionButton3_Click() 
 ListBox1.MultiSelect = fmMultiSelectExtended 
End Sub 
 
Private Sub ToggleButton1_Click() 
 If ToggleButton1.Value = True Then 
 ToggleButton1.Caption = "Plain ListStyle" 
 ListBox1.ListStyle = fmListStylePlain 
 Else 
 ToggleButton1.Caption = "OptionButton " _ 
 & "or CheckBox" 
 ListBox1.ListStyle = fmListStyleOption 
 End If 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]