---
title: Layout event, OldLeft, OldTop, OldHeight, OldWidth properties example
keywords: fm20.chm5225125
f1_keywords:
- fm20.chm5225125
ms.prod: office
ms.assetid: de288917-b1f5-0681-d31f-5847c81b6f29
ms.date: 11/14/2018
localization_priority: Normal
---


# Layout event, OldLeft, OldTop, OldHeight, OldWidth properties example

The following example uses the **[OldLeft and OldTop](oldleft-oldtop-properties.md)** properties and **[OldHeight and OldWidth](oldheight-oldwidth-properties.md)** properties within the **[Layout](layout-event.md)** event to keep a control at its current position and size. 

The user clicks the **[CommandButton](commandbutton-control.md)** labeled **Move ComboBox** to move the control, and then responds to a message box. The user can click the **CommandButton** labeled **Reset ComboBox** to reset the control for another repetition.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- Two **CommandButton** controls named CommandButton1 and CommandButton2.    
- A **[ComboBox](combobox-control.md)** named ComboBox1.
    

```vb
Dim Initialize As Integer 
Dim ComboLeft, ComboTop, ComboWidth, _ 
 ComboHeight As Integer 
 
Private Sub UserForm_Initialize() 
 Initialize = 0 
 CommandButton1.Caption = "Move ComboBox" 
 CommandButton2.Caption = "Reset ComboBox" 
 
 'Information for resetting ComboBox 
 ComboLeft = ComboBox1.Left 
 ComboTop = ComboBox1.Top 
 ComboWidth = ComboBox1.Width 
 ComboHeight = ComboBox1.Height 
End Sub 
 
Private Sub CommandButton1_Click() 
 ComboBox1.Move 0, 0, , , True 
End Sub 
 
Private Sub UserForm_Layout() 
 Dim MyControl As Control 
 Dim MsgBoxResult As Integer 
 'Suppress MsgBox on initial layout event. 
 If Initialize = 0 Then 
 Initialize = 1 
 Exit Sub 
 End If 
 
 MsgBoxResult = MsgBox("In Layout event " _ 
 & "- Continue move?", vbYesNo) 
 If MsgBoxResult = vbNo Then 
 ComboBox1.Move ComboBox1.OldLeft, _ 
 ComboBox1.OldTop, ComboBox1.OldWidth, _ 
 ComboBox1.OldHeight 
 End If 
End Sub 
 
Private Sub CommandButton2_Click() 
 ComboBox1.Move ComboLeft, ComboTop, _ 
 ComboWidth, ComboHeight 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]