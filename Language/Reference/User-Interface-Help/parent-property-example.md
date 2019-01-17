---
title: Parent property example
keywords: fm20.chm5225195
f1_keywords:
- fm20.chm5225195
ms.prod: office
ms.assetid: cad2ce98-5c96-c8b0-4592-f3ffdfdaaed8
ms.date: 11/14/2018
localization_priority: Normal
---


# Parent property example

The following example uses the **[Parent](parent-propertya-microsoft-forms.md)** property to refer to the control or form that contains a specific control.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:

- Two **[Label](label-control.md)** controls named Label1 and Label2.   
- A **[CommandButton](commandbutton-control.md)** named CommandButton1.   
- One or more additional controls of your choice.
    

```vb
Dim MyControl As Object 
Dim MyParent As Object 
Dim ControlsIndex As Integer 
 
Private Sub UserForm_Initialize() 
 ControlsIndex = 0 
 CommandButton1.Caption = "Get Control and Parent" 
 CommandButton1.AutoSize = True 
 CommandButton1.WordWrap = True 
End Sub 
 
Private Sub CommandButton1_Click() 
 'Process Controls collection for UserForm 
 Set MyControl = Controls.Item(ControlsIndex) 
 Set MyParent = MyControl.Parent 
 Label1.Caption = MyControl.Name 
 Label2.Caption = MyParent.Name 
 
 'Prepare index for next control on Userform 
 ControlsIndex = ControlsIndex + 1 
 If ControlsIndex >= Controls.Count Then 
 ControlsIndex = 0 
 End If 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]