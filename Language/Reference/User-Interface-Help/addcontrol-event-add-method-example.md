---
title: AddControl event, Add method example
keywords: fm20.chm5225176
f1_keywords:
- fm20.chm5225176
ms.prod: office
ms.assetid: 6a57bc57-7971-c6b1-72a1-78d5c835b380
ms.date: 11/14/2018
localization_priority: Normal
---


# AddControl event, Add method example

The following example uses the **[Add](add-method-microsoft-forms.md)** method to add a control to a form at run time, and uses the **[AddControl](addcontrol-event.md)** event as confirmation that the control was added.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- A **[CommandButton](commandbutton-control.md)** named CommandButton1.   
- A **[Label](label-control.md)** named Label1.


```vb
Dim Mycmd as Control 
Private Sub CommandButton1_Click() 
 
 Set Mycmd = Controls.Add("MSForms.CommandButton.1") ', CommandButton2, Visible) 
 Mycmd.Left = 18 
 Mycmd.Top = 150 
 Mycmd.Width = 175 
 Mycmd.Height = 20 
 Mycmd.Caption = "This is fun." & Mycmd.Name 
 
End Sub 
 
Private Sub UserForm_AddControl(ByVal Control As _ 
 MSForms.Control) 
 Label1.Caption = "Control was Added." 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]