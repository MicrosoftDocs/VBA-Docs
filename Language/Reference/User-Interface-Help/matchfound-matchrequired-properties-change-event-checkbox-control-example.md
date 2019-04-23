---
title: MatchFound, MatchRequired properties, Change event, CheckBox control example
keywords: fm20.chm5225121
f1_keywords:
- fm20.chm5225121
ms.prod: office
ms.assetid: 10f60293-1b97-faf6-e596-c29489f2439d
ms.date: 11/14/2018
localization_priority: Normal
---


# MatchFound, MatchRequired properties, Change event, CheckBox control example

The following example uses the **[MatchFound](matchfound-property.md)** and **[MatchRequired](matchrequired-property.md)** properties to demonstrate additional character matching for **[ComboBox](combobox-control.md)**. The matching verification occurs in the [Change](change-event.md) event.

In this example, the user specifies whether the text portion of a **ComboBox** must match one of the listed items in the **ComboBox**. The user can specify whether matching is required by using a **[CheckBox](checkbox-control.md)**, and then type into the **ComboBox** to specify an item from its list.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **ComboBox** named ComboBox1.   
- A **CheckBox** named CheckBox1.
    

```vb
Private Sub CheckBox1_Click() 
 If CheckBox1.Value = True Then 
 ComboBox1.MatchRequired = True 
 MsgBox "To move the focus from the " _ 
 & "ComboBox, you must match an entry in " _ 
 & "the list or press ESC." 
 Else 
 ComboBox1.MatchRequired = False 
 MsgBox " To move the focus from the " _ 
 & "ComboBox, just tab to or click " _ 
 & "another control. Matching is optional." 
 End If 
End Sub 
 
Private Sub ComboBox1_Change() 
 If ComboBox1.MatchRequired = True Then 
 'MSForms handles this case automatically 
 Else 
 If ComboBox1.MatchFound = True Then 
 MsgBox "Match Found; matching optional." 
 Else 
 MsgBox "Match not Found; matching " _ 
 & "optional." 
 End If 
 End If 
End Sub 
 
Private Sub UserForm_Initialize() 
Dim i As Integer 
 
For i = 1 To 9 
 ComboBox1.AddItem "Choice " & i 
Next i 
ComboBox1.AddItem "Chocoholic" 
 
CheckBox1.Caption = "MatchRequired" 
CheckBox1.Value = True 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]