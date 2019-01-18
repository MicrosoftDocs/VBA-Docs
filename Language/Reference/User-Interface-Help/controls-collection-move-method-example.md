---
title: Controls collection, Move method example
keywords: fm20.chm5225192
f1_keywords:
- fm20.chm5225192
ms.prod: office
ms.assetid: 14694f03-8d28-9808-b413-96555f0fbc4b
ms.date: 11/14/2018
localization_priority: Normal
---


# Controls collection, Move method example

The following example accesses individual controls from the **[Controls](controls-collection-microsoft-forms.md)** collection by using a **For Each...Next** loop. When the user presses CommandButton1, the other controls are placed in a column along the left edge of the form by using the **[Move](move-method.md)** method.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains a **[CommandButton](commandbutton-control.md)** named CommandButton1 and several other controls.


```vb
Dim CtrlHeight As Single 
Dim CtrlTop As Single 
Dim CtrlGap As Single 
 
Private Sub CommandButton1_Click() 
 Dim MyControl As Control 
 CtrlTop = 5 
 
 For Each MyControl In Controls 
 If MyControl.Name = "CommandButton1" Then 
 'Don't move or resize this control. 
 Else 
 'Move method using named arguments 
 MyControl.Move Top:=CtrlTop, _ 
 Height:=CtrlHeight, Left:=5 
 
 'Move method using unnamed arguments (left, 
 'top, width, height) 
 'MyControl.Move 5, CtrlTop, ,CtrlHeight 
 
 'Calculate top coordinate for next control 
 CtrlTop = CtrlTop + CtrlHeight + CtrlGap 
 End If 
 Next 
 
End Sub
```

<br/>


```vb
Private Sub UserForm_Initialize() 
 CtrlHeight = 20 
 CtrlGap = 5 
 
 CommandButton1.Caption = "Click to move controls" 
 CommandButton1.AutoSize = True 
 CommandButton1.Left = 120 
 CommandButton1.Top = CtrlTop 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]