---
title: ListRows property example
keywords: fm20.chm5225140
f1_keywords:
- fm20.chm5225140
ms.prod: office
ms.assetid: 2161f9e6-a524-b804-79de-acefbeb12180
ms.date: 11/14/2018
localization_priority: Normal
---


# ListRows property example

The following example uses a **[SpinButton](spinbutton-control.md)** to control the number of rows in the drop-down list of a **[ComboBox](combobox-control.md)**. The user changes the value of the **SpinButton**, and then clicks on the drop-down arrow of the **ComboBox** to display the list.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **ComboBox** named ComboBox1.   
- A **SpinButton** named SpinButton1.   
- A **[Label](label-control.md)** named Label1.
    

```vb
Private Sub UserForm_Initialize() 
 Dim i As Integer 
 
 For i = 1 To 20 
 ComboBox1.AddItem "Choice " & (ComboBox1.ListCount + 1) 
 Next i 
 
 SpinButton1.Min = 0 
 SpinButton1.Max = 12 
 SpinButton1.Value = ComboBox1.ListRows 
 Label1.Caption = "ListRows = " & SpinButton1.Value 
End Sub 
 
Private Sub SpinButton1_Change() 
 ComboBox1.ListRows = SpinButton1.Value 
 Label1.Caption = "ListRows = " & SpinButton1.Value 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]