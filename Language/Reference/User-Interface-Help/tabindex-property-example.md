---
title: TabIndex property example
keywords: fm20.chm5225127
f1_keywords:
- fm20.chm5225127
ms.prod: office
ms.assetid: 8329d3f8-0cbd-c520-9659-ff257e4c18d2
ms.date: 11/14/2018
localization_priority: Priority
---


# TabIndex property example

The following example uses the **[TabIndex](tabindex-property.md)** property to display and set the tab order for individual controls. The user can press Tab to reach the next control in the tab order and to display the **TabIndex** of that control. 

The user can also click on a control to display its **TabIndex**. The User can change the **TabIndex** of a control by specifying a new index value in the **[TextBox](textbox-control.md)** and clicking CommandButton3. Changing the **TabIndex** for one control also updates the **TabIndex** for other controls in the **[Frame](frame-control.md)**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **[Label](label-control.md)** named Label1.    
- A **TextBox** named TextBox1.    
- A **Frame** named Frame1.    
- A **TextBox** in the **Frame** named TextBox2.    
- Two **[CommandButton](commandbutton-control.md)** controls in the **Frame** named CommandButton1 and CommandButton2.   
- A **[ScrollBar](scrollbar-control.md)** in the **Frame** named ScrollBar1.    
- A **CommandButton** (not in the **Frame**) named CommandButton3.
    


```vb
Private Sub MoveToFront() 
 Dim i, Temp As Integer 
 
 Temp = Frame1.ActiveControl.TabIndex 
 For i = 0 To Temp - 1 
 Frame1.Controls.Item(i).TabIndex = i + 1 
 Next i 
 
 Frame1.ActiveControl.TabIndex = 0 
 TextBox1.Text = Frame1.ActiveControl.TabIndex 
End Sub 
 
Private Sub CommandButton3_Click() 
 Dim i, Temp As Integer 
 
 If IsNumeric(TextBox1.Text) Then 
 Temp = Val(TextBox1.Text) 
 
 If Temp >= Frame1.Controls.Count Or Temp < 0 
 Then 
 'Entry out of range; move control to front 
 'of tab order 
 MoveToFront 
 ElseIf 
 Temp > Frame1.ActiveControl.TabIndex 
 Then 
 'Move entry down the list 
 For i = Frame1.ActiveControl.TabIndex + _ 
 1 To Temp 
 Frame1.Controls.Item(i).TabIndex = _ 
 i - 1 
 Next i 
 Frame1.ActiveControl.TabIndex = Temp 
 TextBox1.Text = _ 
 Frame1.ActiveControl.TabIndex 
 Else 
 'Move Entry up the list 
 For i = Frame1.ActiveControl.TabIndex - _ 
 1 To Temp 
 Frame1.Controls.Item(i).TabIndex = _ 
 i + 1 
 Next i 
 Frame1.ActiveControl.TabIndex = Temp 
 TextBox1.Text = _ 
 Frame1.ActiveControl.TabIndex 
 End If 
 Else 
 'Text entry; move control to front of tab 
 'order 
 MoveToFront 
 End If 
End Sub 
 
Private Sub UserForm_Initialize() 
 Label1.Caption = "TabIndex" 
 
 Frame1.Controls(0).SetFocus 
 TextBox1.Text = Frame1.ActiveControl.TabIndex 
 
 Frame1.Cycle = fmCycleCurrentForm 
 
 CommandButton3.Caption = "Set TabIndex" 
 CommandButton3.TakeFocusOnClick = False 
End Sub 
 
Private Sub TextBox2_Enter() 
 TextBox1.Text = Frame1.ActiveControl.TabIndex 
End Sub 
 
Private Sub CommandButton1_Enter() 
 TextBox1.Text = Frame1.ActiveControl.TabIndex 
End Sub 
 
Private Sub CommandButton2_Enter() 
 TextBox1.Text = Frame1.ActiveControl.TabIndex 
End Sub 
 
Private Sub ScrollBar1_Enter() 
 TextBox1.Text = Frame1.ActiveControl.TabIndex 
End Sub
```


